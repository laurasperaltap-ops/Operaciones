#!/usr/bin/env python3
"""Normalizacion y geocodificacion de direcciones (Origen/Destino).

Resuelve cada direccion contra Google Maps y extrae Estado / Municipio /
Parroquia / Ciudad. Solo acepta un resultado cuando el geocoder aporta evidencia
al nivel de la direccion: si la API resuelve unicamente a la ciudad o al estado,
la fila se marca "Sin match" en vez de rellenarse con la ubicacion generica.

Diseño:
  - Cache SQLite con la respuesta cruda de la API, indexada por (api, consulta).
    Reprocesar con otra heuristica no vuelve a pagar llamadas ya hechas.
  - Checkpoint JSONL por `ID cotización`. Al reanudar se saltan las filas ya
    resueltas sin volver a tocarlas.
  - Limite de tasa explicito y backoff exponencial con jitter ante 429,
    OVER_QUERY_LIMIT y errores 5xx.

Uso:
    python tools/geocodificar.py --anio 2026 --sublote 1
    python tools/geocodificar.py --anio 2026            # todos los sublotes
    python tools/geocodificar.py --anio 2026 --exportar salidas/2026.xlsx
"""

from __future__ import annotations

import argparse
import json
import os
import random
import re
import sqlite3
import sys
import time
import unicodedata
from dataclasses import dataclass, asdict, field
from pathlib import Path

import openpyxl
import pandas as pd
import requests

RAIZ = Path(__file__).resolve().parent.parent
FUENTE_POR_DEFECTO = RAIZ / "datos" / "Base_rentas_analisis_DATA_DIRECCIONES.xlsx"
DIR_CHECKPOINTS = RAIZ / "checkpoints"
DIR_SALIDAS = RAIZ / "salidas"
CACHE = RAIZ / "checkpoints" / "cache_geocoding.sqlite"

COL_ID = "ID cotización"
COL_ANIO = "Año"
COL_CIUDAD = "Ciudad"
COL_ORIGEN = "Origen "  # el nombre en la fuente trae un espacio final
COL_DESTINO = "Destino"

PAIS = "Venezuela"
CODIGO_PAIS = "VE"

# Tipos que NO constituyen una direccion: si el mejor resultado se queda en este
# nivel, el geocoder no encontro el lugar y devolvio el contenedor administrativo.
TIPOS_GENERICOS = {
    "locality", "political", "administrative_area_level_1",
    "administrative_area_level_2", "administrative_area_level_3",
    "country", "postal_code",
}

# Tipos que sí acreditan una direccion concreta.
TIPOS_PRECISOS = {
    "street_address", "premise", "subpremise", "establishment",
    "point_of_interest", "airport", "route", "intersection",
    "transit_station", "bus_station", "lodging", "neighborhood",
    "sublocality", "sublocality_level_1", "shopping_mall",
}


def normalizar(texto: object) -> str:
    """Minusculas sin acentos ni espacios extremos."""
    if texto is None:
        return ""
    base = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in base if not unicodedata.combining(c)).strip().lower()


def esta_vacio(valor: object) -> bool:
    return valor is None or pd.isna(valor) or not str(valor).strip()


# --------------------------------------------------------------------------- #
# Cache
# --------------------------------------------------------------------------- #

class CacheRespuestas:
    """Almacen persistente de respuestas crudas de la API.

    Guardar el JSON completo permite reinterpretar los resultados con otra
    heuristica sin repetir (ni pagar) las llamadas.
    """

    def __init__(self, ruta: Path):
        ruta.parent.mkdir(parents=True, exist_ok=True)
        self.con = sqlite3.connect(ruta)
        self.con.execute(
            "CREATE TABLE IF NOT EXISTS respuestas ("
            " api TEXT NOT NULL, consulta TEXT NOT NULL, json TEXT NOT NULL,"
            " obtenido_en TEXT NOT NULL, PRIMARY KEY (api, consulta))"
        )
        self.con.commit()
        self.aciertos = 0
        self.fallos = 0

    def obtener(self, api: str, consulta: str) -> dict | None:
        fila = self.con.execute(
            "SELECT json FROM respuestas WHERE api=? AND consulta=?", (api, consulta)
        ).fetchone()
        if fila:
            self.aciertos += 1
            return json.loads(fila[0])
        self.fallos += 1
        return None

    def guardar(self, api: str, consulta: str, payload: dict) -> None:
        self.con.execute(
            "INSERT OR REPLACE INTO respuestas VALUES (?,?,?,datetime('now'))",
            (api, consulta, json.dumps(payload, ensure_ascii=False)),
        )
        self.con.commit()


# --------------------------------------------------------------------------- #
# Cliente HTTP
# --------------------------------------------------------------------------- #

class LimiteExcedido(RuntimeError):
    """La cuota diaria se agoto: no tiene sentido seguir reintentando."""


class ClienteMaps:
    """Cliente con limite de tasa y reintentos ante saturacion."""

    ESTADOS_REINTENTABLES = {"OVER_QUERY_LIMIT", "UNKNOWN_ERROR"}
    ESTADOS_SIN_RESULTADO = {"ZERO_RESULTS", "NOT_FOUND"}

    def __init__(self, clave: str, cache: CacheRespuestas, qps: float = 10.0,
                 max_reintentos: int = 5):
        self.clave = clave
        self.cache = cache
        self.intervalo = 1.0 / qps if qps > 0 else 0.0
        self.max_reintentos = max_reintentos
        self.ultima_llamada = 0.0
        self.llamadas_reales = 0
        self.sesion = requests.Session()

    def _esperar_turno(self) -> None:
        transcurrido = time.monotonic() - self.ultima_llamada
        if transcurrido < self.intervalo:
            time.sleep(self.intervalo - transcurrido)
        self.ultima_llamada = time.monotonic()

    def pedir(self, api: str, parametros: dict) -> dict:
        """Llama a la API respetando cache, limite de tasa y backoff."""
        clave_cache = json.dumps(parametros, sort_keys=True, ensure_ascii=False)
        guardado = self.cache.obtener(api, clave_cache)
        if guardado is not None:
            return guardado

        url = {
            "geocode": "https://maps.googleapis.com/maps/api/geocode/json",
            "places": "https://maps.googleapis.com/maps/api/place/textsearch/json",
        }[api]

        for intento in range(self.max_reintentos):
            self._esperar_turno()
            try:
                respuesta = self.sesion.get(
                    url, params={**parametros, "key": self.clave}, timeout=30
                )
                self.llamadas_reales += 1
            except requests.RequestException as error:
                if intento == self.max_reintentos - 1:
                    raise
                self._dormir(intento, f"red: {error}")
                continue

            if respuesta.status_code == 429 or respuesta.status_code >= 500:
                if intento == self.max_reintentos - 1:
                    raise RuntimeError(f"HTTP {respuesta.status_code} tras {intento+1} intentos")
                self._dormir(intento, f"HTTP {respuesta.status_code}")
                continue

            datos = respuesta.json()
            estado = datos.get("status")

            # Cuota agotada: reintentar solo empeora las cosas.
            if estado == "OVER_DAILY_LIMIT" or (
                estado == "REQUEST_DENIED" and "quota" in str(datos.get("error_message", "")).lower()
            ):
                raise LimiteExcedido(datos.get("error_message", estado))

            if estado in self.ESTADOS_REINTENTABLES and intento < self.max_reintentos - 1:
                self._dormir(intento, estado)
                continue

            if estado == "REQUEST_DENIED":
                raise RuntimeError(f"REQUEST_DENIED: {datos.get('error_message')}")

            self.cache.guardar(api, clave_cache, datos)
            return datos

        raise RuntimeError(f"Sin respuesta util de {api} tras {self.max_reintentos} intentos")

    def _dormir(self, intento: int, motivo: str) -> None:
        espera = (2 ** intento) + random.uniform(0, 1)
        print(f"    [backoff] {motivo} — reintento en {espera:.1f}s", file=sys.stderr)
        time.sleep(espera)


# --------------------------------------------------------------------------- #
# Interpretacion de resultados
# --------------------------------------------------------------------------- #

@dataclass
class Resolucion:
    """Lo que se pudo determinar de una direccion, con su procedencia."""
    estado_match: str = "Sin match"
    precision: str = "sin_match"
    fuente: str = ""
    consulta: str = ""
    estado: str = "Sin match"
    municipio: str = "Sin match"
    parroquia: str = "Sin match"
    ciudad: str = "Sin match"
    direccion_formateada: str = "Sin match"
    lat: float | None = None
    lng: float | None = None
    place_id: str = ""
    tipos: str = ""
    motivo: str = ""


def _componente(resultado: dict, *tipos: str) -> str | None:
    """Primer componente que case con alguno de los tipos, por orden de preferencia."""
    for tipo in tipos:
        for componente in resultado.get("address_components", []):
            if tipo in componente["types"]:
                return componente["long_name"]
    return None


# Palabras que no aportan ubicacion: si al quitar los nombres de lugar resueltos
# solo queda esto, el texto no pedia nada mas especifico que la ciudad.
PALABRAS_VACIAS = {
    "a", "al", "ante", "con", "de", "del", "desde", "donde", "e", "el", "en",
    "entre", "hacia", "hasta", "la", "las", "lo", "los", "o", "para", "por",
    "segun", "sin", "sobre", "un", "una", "y",
    "disposicion", "disponible", "recorrido", "traslado", "viaje", "ruta",
    "varios", "puntos", "destino", "origen", "zona", "area", "ciudad",
    "estado", "municipio", "sector", "via", "s", "n",
}

# Abreviaturas de la fuente que el geocoder resuelve bien pero que no coinciden
# textualmente con el nombre devuelto.
ALIAS_LUGARES = {
    "ccs": "caracas", "mcbo": "maracaibo", "mcb": "maracaibo", "mbo": "maracaibo",
    "bqto": "barquisimeto", "pzo": "puerto ordaz", "plc": "puerto la cruz",
    "pmv": "porlamar", "vln": "valencia",
}

# Localidades que, para esta base, quedan dentro del area que nombra la columna
# Ciudad. Sin esta tabla, una fila declarada 'Caracas' que resuelve en Maiquetia
# o Catia La Mar se marcaria como conflicto siendo el litoral de la misma area
# metropolitana, y una declarada 'Margarita' chocaria con todo Nueva Esparta.
EQUIVALENCIAS_AREA = {
    "caracas": {"caracas", "distrito capital", "libertador", "miranda", "chacao",
                "baruta", "sucre", "el hatillo", "la guaira", "vargas", "maiquetia",
                "catia la mar", "caraballeda", "macuto", "naiguata", "camuri grande"},
    "la guaira": {"la guaira", "vargas", "maiquetia", "catia la mar", "caraballeda",
                  "macuto", "naiguata", "camuri grande", "caracas", "distrito capital"},
    "maracaibo": {"maracaibo", "san francisco", "zulia"},
    "margarita": {"margarita", "nueva esparta", "porlamar", "pampatar", "la asuncion",
                  "juan griego", "el yaque", "punta de piedras", "maneiro", "mariño", "diaz"},
    "valencia": {"valencia", "carabobo", "naguanagua", "san diego", "los guayos"},
    "barquisimeto": {"barquisimeto", "lara", "iribarren", "cabudare", "palavecino"},
    "maracay": {"maracay", "aragua", "girardot", "linares alcantara"},
    "los teques": {"los teques", "miranda", "guaicaipuro", "carrizal",
                   "san antonio de los altos"},
    "barcelona": {"barcelona", "anzoategui", "puerto la cruz", "lecheria", "bolivar"},
}

# Un plus-code sin calle ni establecimiento delante indica que la API no hallo
# nada con nombre y devolvio una coordenada del area.
PATRON_PLUSCODE = re.compile(r"^[23456789CFGHJMPQRVWX]{4,8}\+[23456789CFGHJMPQRVWX]{2,3}\b",
                             re.IGNORECASE)


def clasificar(resultado: dict) -> str:
    """Precision del resultado segun sus tipos y si fue coincidencia parcial."""
    tipos = set(resultado.get("types", []))
    if tipos & TIPOS_PRECISOS:
        return "aproximada" if resultado.get("partial_match") else "exacta"
    if "country" in tipos:
        # Resolvio al pais entero: nunca es una direccion.
        return "solo_pais"
    if tipos & TIPOS_GENERICOS:
        return "solo_ciudad"
    return "desconocida"


def texto_pedia_solo_la_ciudad(texto_entrada: str, resolucion: Resolucion) -> bool:
    """El texto no nombraba nada mas especifico que el lugar ya resuelto.

    Distingue las dos caras del resultado a nivel de localidad. Si el texto dice
    'Caracas' o 'A disposición en CCS', resolver a la ciudad es la respuesta
    completa. Si dice 'Distribuidora Maracaibo Sur.', la ciudad es lo que quedo
    tras no encontrar el negocio, y sobra 'distribuidora sur' como prueba de que
    se pedia algo mas preciso.
    """
    objetivo = normalizar(texto_entrada)
    for abreviatura, expandido in ALIAS_LUGARES.items():
        objetivo = re.sub(rf"\b{abreviatura}\b", expandido, objetivo)

    if not re.findall(r"[a-z0-9]+", objetivo):
        # Un guion o una celda con simbolos no pidio la ciudad: no pidio nada.
        return False

    # La ciudad tiene que salir del texto, no del sesgo de la consulta. Sin esta
    # comprobacion, un texto como 'DISPOSICIÓN' se daria por resuelto en la
    # ciudad declarada sin que nada en el texto la mencione.
    menciona_el_lugar = False
    for valor in (resolucion.ciudad, resolucion.estado, resolucion.municipio,
                  resolucion.parroquia):
        if not valor or valor == "Sin match":
            continue
        patron = rf"\b{re.escape(normalizar(valor))}\b"
        if re.search(patron, objetivo):
            menciona_el_lugar = True
            objetivo = re.sub(patron, " ", objetivo)
    if not menciona_el_lugar:
        return False

    restante = [t for t in re.findall(r"[a-z0-9]+", objetivo)
                if t not in PALABRAS_VACIAS]
    return not restante


def texto_corrobora(texto_entrada: str, resolucion: Resolucion) -> bool:
    """El texto original menciona alguno de los lugares que devolvio el geocoder.

    Es la evidencia que separa una correccion legitima de una deriva. Si el
    usuario escribio 'urb. Carabobo' y la API responde estado Carabobo, el texto
    respalda el resultado aunque contradiga la columna Ciudad. Si escribio
    'escuela rafael Rangel' y la API responde Trujillo, nada lo respalda.
    """
    objetivo = normalizar(texto_entrada)
    for valor in (resolucion.estado, resolucion.municipio, resolucion.ciudad,
                  resolucion.parroquia):
        if not valor or valor == "Sin match":
            continue
        aguja = normalizar(valor)
        # Frontera de palabra: evita que 'Lara' case dentro de otra palabra.
        if aguja and re.search(rf"\b{re.escape(aguja)}\b", objetivo):
            return True
    return False


def interpretar(resultado: dict, fuente: str, consulta: str,
                texto_entrada: str = "", ciudad_declarada: object = None) -> Resolucion:
    """Convierte un resultado de la API en una Resolucion, o la rechaza."""
    pais = _componente(resultado, "country")
    if pais is None:
        # Sin address_components no hay forma de acreditar el pais ni de extraer
        # estado/municipio/parroquia. Ocurre con respuestas que no se expandieron.
        return Resolucion(precision="sin_componentes", fuente=fuente, consulta=consulta,
                          motivo="la respuesta no trae address_components")
    if normalizar(pais) != normalizar(PAIS):
        return Resolucion(precision="fuera_venezuela", fuente=fuente, consulta=consulta,
                          motivo=f"pais={pais!r}")

    precision = clasificar(resultado)
    if precision in ("solo_pais", "desconocida"):
        return Resolucion(precision=precision, fuente=fuente, consulta=consulta,
                          motivo=f"tipos={','.join(resultado.get('types', []))}")

    ubicacion = resultado.get("geometry", {}).get("location", {})
    resolucion = Resolucion(
        estado_match="Match",
        precision=precision,
        fuente=fuente,
        consulta=consulta,
        estado=_componente(resultado, "administrative_area_level_1") or "Sin match",
        municipio=_componente(resultado, "administrative_area_level_2") or "Sin match",
        parroquia=_componente(resultado, "administrative_area_level_3",
                              "sublocality_level_1", "sublocality", "neighborhood") or "Sin match",
        ciudad=_componente(resultado, "locality", "administrative_area_level_2") or "Sin match",
        direccion_formateada=resultado.get("formatted_address", "Sin match"),
        lat=ubicacion.get("lat"),
        lng=ubicacion.get("lng"),
        place_id=resultado.get("place_id", ""),
        tipos=",".join(resultado.get("types", [])),
    )

    if precision == "solo_ciudad":
        # La API se quedo en la localidad. Es la respuesta correcta solo si el
        # texto tampoco pedia nada mas fino que eso.
        if texto_pedia_solo_la_ciudad(texto_entrada, resolucion):
            resolucion.precision = "nivel_ciudad"
        else:
            return Resolucion(
                precision="solo_ciudad", fuente=fuente, consulta=consulta,
                motivo=f"no hallo el lugar pedido; devolvio {resolucion.ciudad!r} "
                       f"(tipos={','.join(resultado.get('types', []))})")

    corroborado = texto_corrobora(texto_entrada, resolucion)

    # El resultado cae fuera del area declarada y nada en el texto lo respalda:
    # se conserva lo que dijo el geocoder, pero no se da por bueno.
    if (not esta_vacio(ciudad_declarada)
            and hay_conflicto(ciudad_declarada, resolucion)
            and not corroborado):
        resolucion.estado_match = "Revisar"
        resolucion.precision = "conflicto_geografico"
        resolucion.motivo = (f"resolvio en {resolucion.ciudad}/{resolucion.estado} "
                             f"pero la fuente declara {str(ciudad_declarada).strip()!r} "
                             f"y el texto no lo respalda")
        return resolucion

    # Coordenada sin nombre: la API no encontro el lugar y devolvio un punto del
    # area. El plus-code por si solo no alcanza como sintoma, porque Google lo
    # antepone tambien a direcciones que si traen calle; lo que delata al caso
    # vacio es que ademas no venga ningun componente de via o establecimiento.
    tiene_via = any(
        {"route", "establishment", "premise", "subpremise", "point_of_interest",
         "airport", "natural_feature", "park"} & set(componente["types"])
        for componente in resultado.get("address_components", [])
    )
    if (resolucion.precision != "nivel_ciudad"
            and PATRON_PLUSCODE.match(resolucion.direccion_formateada.strip())
            and not tiene_via
            and not corroborado):
        resolucion.estado_match = "Revisar"
        resolucion.precision = "coordenada_sin_nombre"
        resolucion.motivo = (f"la API devolvio solo un plus-code en {resolucion.ciudad} "
                             f"sin calle ni establecimiento; el texto no lo respalda")

    return resolucion


def expandir_place(cliente: ClienteMaps, place_id: str) -> dict | None:
    """Reconsulta un place_id al geocoder para obtener sus address_components."""
    if not place_id:
        return None
    datos = cliente.pedir("geocode", {"place_id": place_id, "language": "es"})
    if datos.get("status") == "OK" and datos.get("results"):
        return datos["results"][0]
    return None


def resolver(cliente: ClienteMaps, texto: str, ciudad: str) -> Resolucion:
    """Cascada de busqueda: de la consulta mas contextualizada a la mas amplia.

    Cada escalon solo se intenta si el anterior no aporto evidencia a nivel de
    direccion. Devuelve la primera Resolucion aceptada, o la ultima rechazada
    para conservar el motivo del descarte.
    """
    if esta_vacio(texto):
        return Resolucion(precision="texto_vacio", motivo="Origen/Destino vacio en la fuente")

    texto = str(texto).strip()
    tiene_ciudad = not esta_vacio(ciudad)
    ciudad_txt = str(ciudad).strip() if tiene_ciudad else ""

    # Orden deliberado: primero lo mas contextualizado. La consulta sin ciudad va
    # ultima porque, al soltar el unico ancla geografica, es la que puede derivar
    # a otro estado cuando el texto no dice donde queda el lugar.
    intentos: list[tuple[str, dict]] = []
    if tiene_ciudad:
        intentos.append(("geocode", {
            "address": f"{texto}, {ciudad_txt}, {PAIS}",
            "components": f"country:{CODIGO_PAIS}", "region": "ve", "language": "es",
        }))
        # Places entiende nombres de negocios y puntos de interes que el geocoder no.
        intentos.append(("places", {
            "query": f"{texto}, {ciudad_txt}, {PAIS}", "region": "ve", "language": "es",
        }))
    else:
        intentos.append(("places", {
            "query": f"{texto}, {PAIS}", "region": "ve", "language": "es",
        }))
    intentos.append(("geocode", {
        "address": f"{texto}, {PAIS}",
        "components": f"country:{CODIGO_PAIS}", "region": "ve", "language": "es",
    }))

    ultimo_rechazo = Resolucion(motivo="sin resultados en ningun escalon")

    for api, parametros in intentos:
        datos = cliente.pedir(api, parametros)
        consulta = parametros.get("address") or parametros.get("query", "")

        if datos.get("status") in ClienteMaps.ESTADOS_SIN_RESULTADO:
            ultimo_rechazo = Resolucion(precision="sin_match", fuente=api, consulta=consulta,
                                        motivo=datos["status"])
            continue
        if datos.get("status") != "OK" or not datos.get("results"):
            ultimo_rechazo = Resolucion(precision="sin_match", fuente=api, consulta=consulta,
                                        motivo=str(datos.get("status")))
            continue

        crudo = datos["results"][0]
        if api == "places":
            # Text Search no devuelve address_components; sin ellos no hay estado
            # ni municipio que extraer. Se reconsulta por place_id al geocoder,
            # que sí los trae. La llamada extra queda cacheada.
            crudo = expandir_place(cliente, crudo.get("place_id", "")) or crudo

        resolucion = interpretar(crudo, api, consulta, texto, ciudad)
        if resolucion.estado_match == "Match":
            return resolucion
        # "Revisar" trae datos utiles y gana a un rechazo pelado, pero se sigue
        # buscando por si un escalon posterior devuelve un match limpio.
        if resolucion.estado_match == "Revisar" or ultimo_rechazo.estado_match != "Revisar":
            ultimo_rechazo = resolucion

    return ultimo_rechazo


# --------------------------------------------------------------------------- #
# Proceso por lotes con checkpoint
# --------------------------------------------------------------------------- #

@dataclass
class RegistroFila:
    id_cotizacion: int
    ciudad_declarada: str
    origen: dict = field(default_factory=dict)
    destino: dict = field(default_factory=dict)
    conflicto_ciudad: bool = False


def cargar_procesados(ruta: Path) -> dict[int, dict]:
    """IDs ya resueltos en corridas previas, para no reprocesarlos."""
    if not ruta.exists():
        return {}
    procesados: dict[int, dict] = {}
    with ruta.open(encoding="utf-8") as archivo:
        for linea in archivo:
            linea = linea.strip()
            if not linea:
                continue
            try:
                registro = json.loads(linea)
            except json.JSONDecodeError:
                # Una corrida interrumpida a mitad de escritura deja una linea
                # truncada al final; se descarta y se reprocesa esa fila.
                print(f"[aviso] linea de checkpoint ilegible, se ignora", file=sys.stderr)
                continue
            procesados[registro["id_cotizacion"]] = registro
    return procesados


def hay_conflicto(declarada: object, resolucion: "Resolucion") -> bool:
    """El lugar resuelto queda fuera del area que nombra la columna Ciudad.

    Compara contra todos los niveles administrativos devueltos, no solo la
    localidad: una fila de 'Caracas' puede resolver con localidad 'Chacao' y
    estado 'Miranda' sin que eso sea un conflicto.
    """
    if esta_vacio(declarada) or resolucion.ciudad in ("", "Sin match"):
        return False

    declarada_norm = normalizar(declarada)
    candidatos = set()
    for valor in (resolucion.ciudad, resolucion.municipio, resolucion.estado,
                  resolucion.parroquia):
        if valor and valor != "Sin match":
            limpio = normalizar(valor)
            candidatos.add(limpio)
            # 'Municipio Autonomo Valencia' tiene que poder casar con 'Valencia'.
            candidatos.add(re.sub(r"^(municipio|parroquia)\s+(autonomo\s+)?(urbana\s+)?",
                                  "", limpio))

    if declarada_norm in candidatos:
        return False
    return not (candidatos & EQUIVALENCIAS_AREA.get(declarada_norm, set()))


def procesar(marco: pd.DataFrame, cliente: ClienteMaps, ruta_checkpoint: Path,
             cada: int, sublote: int | None, tamano_sublote: int) -> dict:
    procesados = cargar_procesados(ruta_checkpoint)
    if procesados:
        print(f"Reanudando: {len(procesados)} filas ya resueltas en {ruta_checkpoint.name}")

    pendientes = marco[~marco[COL_ID].isin(procesados.keys())]

    if sublote is not None:
        inicio = (sublote - 1) * tamano_sublote
        pendientes = pendientes.iloc[inicio:inicio + tamano_sublote]
        print(f"Sublote {sublote}: filas {inicio}-{inicio + len(pendientes)} del pendiente")

    print(f"A procesar en esta corrida: {len(pendientes)} filas\n")
    ruta_checkpoint.parent.mkdir(parents=True, exist_ok=True)

    sin_grabar = 0
    archivo = ruta_checkpoint.open("a", encoding="utf-8")
    try:
        for contador, (_, fila) in enumerate(pendientes.iterrows(), start=1):
            identificador = int(fila[COL_ID])
            ciudad = fila.get(COL_CIUDAD)

            origen = resolver(cliente, fila.get(COL_ORIGEN), ciudad)
            destino = resolver(cliente, fila.get(COL_DESTINO), ciudad)

            registro = RegistroFila(
                id_cotizacion=identificador,
                ciudad_declarada="" if esta_vacio(ciudad) else str(ciudad).strip(),
                origen=asdict(origen),
                destino=asdict(destino),
                conflicto_ciudad=(hay_conflicto(ciudad, origen)
                                  or hay_conflicto(ciudad, destino)),
            )
            archivo.write(json.dumps(asdict(registro), ensure_ascii=False) + "\n")
            sin_grabar += 1

            if sin_grabar >= cada:
                archivo.flush()
                os.fsync(archivo.fileno())
                sin_grabar = 0
                print(f"  [checkpoint] {contador}/{len(pendientes)} — "
                      f"llamadas reales={cliente.llamadas_reales} "
                      f"cache={cliente.cache.aciertos}")

            if contador % 25 == 0:
                print(f"  {contador}/{len(pendientes)} filas "
                      f"(ultimo id={identificador}, origen={origen.precision}, "
                      f"destino={destino.precision})")
    except LimiteExcedido as error:
        print(f"\n[CUOTA AGOTADA] {error}", file=sys.stderr)
        print(f"Checkpoint guardado en {ruta_checkpoint}. Reanudar con el mismo comando.",
              file=sys.stderr)
    except KeyboardInterrupt:
        print("\n[interrumpido] guardando checkpoint...", file=sys.stderr)
    finally:
        archivo.flush()
        os.fsync(archivo.fileno())
        archivo.close()

    return cargar_procesados(ruta_checkpoint)


# --------------------------------------------------------------------------- #
# Salida
# --------------------------------------------------------------------------- #

CAMPOS_SALIDA = [
    ("estado_match", "Estado match"), ("precision", "Precisión"), ("fuente", "Fuente"),
    ("estado", "Estado"), ("municipio", "Municipio"), ("parroquia", "Parroquia"),
    ("ciudad", "Ciudad geocodificada"), ("direccion_formateada", "Dirección normalizada"),
    ("lat", "Latitud"), ("lng", "Longitud"), ("consulta", "Consulta usada"),
    ("place_id", "Place ID"), ("motivo", "Motivo sin match"),
]


def construir_salida(marco: pd.DataFrame, procesados: dict[int, dict]) -> pd.DataFrame:
    """Adosa las columnas normalizadas al marco original, sin tocar lo existente."""
    salida = marco.copy()
    for prefijo, lado in (("Origen", "origen"), ("Destino", "destino")):
        for campo, etiqueta in CAMPOS_SALIDA:
            salida[f"{prefijo} · {etiqueta}"] = salida[COL_ID].map(
                lambda i: (procesados.get(int(i), {}).get(lado) or {}).get(campo)
            )
    salida["Conflicto ciudad declarada"] = salida[COL_ID].map(
        lambda i: procesados.get(int(i), {}).get("conflicto_ciudad")
    )
    return salida


def escribir_excel(marco: pd.DataFrame, ruta: Path) -> None:
    """Escribe el libro preservando los literales de error de la fuente.

    openpyxl tipa como celda de error cualquier texto que sea un codigo de Excel
    ('#VALUE!', '#REF!'), y al releer vuelve como nulo. Como esos literales
    forman parte de los datos originales, se reescriben como texto para que la
    salida no altere ninguna columna de entrada.
    """
    from openpyxl.cell.cell import TYPE_STRING

    ruta.parent.mkdir(parents=True, exist_ok=True)
    marco.to_excel(ruta, index=False)

    libro = openpyxl.load_workbook(ruta)
    reparadas = 0
    for hoja in libro.worksheets:
        for fila in hoja.iter_rows():
            for celda in fila:
                if celda.data_type == "e" and celda.value is not None:
                    celda.data_type = TYPE_STRING
                    reparadas += 1
    if reparadas:
        libro.save(ruta)
        print(f"  ({reparadas} literales de error preservados como texto)")


def resumen(procesados: dict[int, dict], total_filas: int, cliente: ClienteMaps) -> str:
    lineas = ["=" * 68, "RESUMEN DE CORRIDA", "=" * 68,
              f"Filas del lote          : {total_filas}",
              f"Filas resueltas         : {len(procesados)}", ""]
    for lado, titulo in (("origen", "ORIGEN"), ("destino", "DESTINO")):
        conteo: dict[str, int] = {}
        for registro in procesados.values():
            clave = (registro.get(lado) or {}).get("precision", "?")
            conteo[clave] = conteo.get(clave, 0) + 1
        con_match = sum(1 for r in procesados.values()
                        if (r.get(lado) or {}).get("estado_match") == "Match")
        pct = con_match / len(procesados) * 100 if procesados else 0
        lineas.append(f"{titulo}: {con_match}/{len(procesados)} con match ({pct:.1f}%)")
        for clave, valor in sorted(conteo.items(), key=lambda x: -x[1]):
            lineas.append(f"    {clave:<18}: {valor}")
        lineas.append("")
    conflictos = sum(1 for r in procesados.values() if r.get("conflicto_ciudad"))
    lineas += [f"Filas con ciudad en conflicto: {conflictos}",
               f"Llamadas reales a la API     : {cliente.llamadas_reales}",
               f"Aciertos de cache            : {cliente.cache.aciertos}", "=" * 68]
    return "\n".join(lineas)


def main() -> None:
    analizador = argparse.ArgumentParser(description=__doc__,
                                         formatter_class=argparse.RawDescriptionHelpFormatter)
    analizador.add_argument("--fuente", type=Path, default=FUENTE_POR_DEFECTO)
    analizador.add_argument("--anio", type=int, required=True)
    analizador.add_argument("--sublote", type=int, default=None,
                            help="Procesa solo el sublote N de los pendientes")
    analizador.add_argument("--tamano-sublote", type=int, default=200)
    analizador.add_argument("--checkpoint-cada", type=int, default=25)
    analizador.add_argument("--qps", type=float, default=10.0)
    analizador.add_argument("--exportar", type=Path, default=None)
    analizador.add_argument("--solo-resumen", action="store_true",
                            help="No llama a la API: resume el checkpoint existente")
    args = analizador.parse_args()

    clave = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if not clave:
        raise SystemExit("Falta GOOGLE_MAPS_API_KEY en el entorno.")

    marco = pd.read_excel(args.fuente, sheet_name=0).dropna(how="all")
    lote = marco[pd.to_numeric(marco[COL_ANIO], errors="coerce") == args.anio].copy()
    print(f"Lote {args.anio}: {len(lote)} filas de {len(marco)} totales")

    ruta_checkpoint = DIR_CHECKPOINTS / f"{args.anio}.jsonl"
    cache = CacheRespuestas(CACHE)
    cliente = ClienteMaps(clave, cache, qps=args.qps)

    if args.solo_resumen:
        procesados = cargar_procesados(ruta_checkpoint)
    else:
        procesados = procesar(lote, cliente, ruta_checkpoint,
                              args.checkpoint_cada, args.sublote, args.tamano_sublote)

    print()
    print(resumen(procesados, len(lote), cliente))

    if args.exportar:
        escribir_excel(construir_salida(lote, procesados), args.exportar)
        print(f"\nExportado a {args.exportar}")


if __name__ == "__main__":
    main()
