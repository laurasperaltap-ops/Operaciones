#!/usr/bin/env python3
"""Distancia y duracion en auto para las filas con ambos extremos resueltos.

Lee el entregable ya geocodificado y agrega dos columnas al final de la hoja
'Direcciones 2026'. No vuelve a geocodificar nada: toma las coordenadas y los
place_id que esa hoja ya trae. Las demas hojas y las columnas existentes quedan
intactas, porque el libro se edita en sitio con openpyxl en vez de reescribirse.

Uso:
    python tools/calcular_distancias.py --libro salidas/Entregable_direcciones_2026.xlsx
    python tools/calcular_distancias.py --libro <ruta> --sublote 150
    python tools/calcular_distancias.py --libro <ruta> --solo-reporte
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from geocodificar import CACHE, CacheRespuestas, ClienteMaps, LimiteExcedido  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
CHECKPOINT = RAIZ / "checkpoints" / "distancias.jsonl"

COL_ID = "ID cotización"
COL_DISTANCIA = "Distancia (KM)"
COL_DURACION = "Duración estimada (HH:MM)"
COL_TIPO = "Distancia · Tipo de cálculo"

# Precisiones que no apuntan a una direccion concreta: su coordenada ya es el
# centroide que Google usa para la localidad, asi que el trayecto se mide entre
# centroides y la fila queda marcada como aproximada.
PRECISIONES_CENTROIDE = {"nivel_ciudad"}


def formatear_duracion(segundos: int) -> str:
    horas, resto = divmod(int(segundos), 3600)
    return f"{horas:02d}:{resto // 60:02d}"


def punto_de(fila: dict, prefijo: str) -> str | None:
    """Referencia que se manda a la API: place_id si esta, si no lat,lng.

    El place_id ancla mejor sobre establecimientos que una coordenada suelta,
    que la API tiene que enganchar a la calle mas cercana.
    """
    place_id = fila.get(f"{prefijo} · Place ID")
    if place_id and str(place_id).strip() and str(place_id).strip().lower() != "nan":
        return f"place_id:{str(place_id).strip()}"
    lat = fila.get(f"{prefijo} · Latitud")
    lng = fila.get(f"{prefijo} · Longitud")
    if lat is None or lng is None or lat == "" or lng == "":
        return None
    try:
        return f"{float(lat)},{float(lng)}"
    except (TypeError, ValueError):
        return None


def tipo_de_calculo(fila: dict) -> str:
    origen = str(fila.get("Origen · Precisión", "")) in PRECISIONES_CENTROIDE
    destino = str(fila.get("Destino · Precisión", "")) in PRECISIONES_CENTROIDE
    if origen and destino:
        return "Aproximado · centroide en ambos extremos"
    if origen:
        return "Aproximado · centroide en origen"
    if destino:
        return "Aproximado · centroide en destino"
    return "Punto a punto"


def hoja_de_datos(libro: openpyxl.Workbook) -> str:
    """La hoja con la tabla completa, cuyo nombre lleva el anio del lote."""
    for nombre in libro.sheetnames:
        if nombre.startswith("Direcciones "):
            return nombre
    raise SystemExit(f"No hay hoja 'Direcciones <anio>'. Hojas: {libro.sheetnames}")


def leer_hoja(libro_ruta: Path) -> tuple[openpyxl.Workbook, list[str], list[dict]]:
    libro = openpyxl.load_workbook(libro_ruta)
    hoja = libro[hoja_de_datos(libro)]
    encabezados = [c.value for c in hoja[1]]
    filas = []
    for numero, fila in enumerate(hoja.iter_rows(min_row=2, values_only=True), start=2):
        registro = dict(zip(encabezados, fila))
        registro["__fila__"] = numero
        filas.append(registro)
    return libro, encabezados, filas


def candidatas(filas: list[dict]) -> list[dict]:
    """Filas con ambos extremos en Match: las unicas medibles por ahora."""
    return [f for f in filas
            if str(f.get("Origen · Estado match", "")).strip() == "Match"
            and str(f.get("Destino · Estado match", "")).strip() == "Match"]


def cargar_checkpoint() -> dict[int, dict]:
    if not CHECKPOINT.exists():
        return {}
    hechas = {}
    with CHECKPOINT.open(encoding="utf-8") as archivo:
        for linea in archivo:
            linea = linea.strip()
            if not linea:
                continue
            try:
                registro = json.loads(linea)
            except json.JSONDecodeError:
                print("[aviso] linea de checkpoint ilegible, se ignora", file=sys.stderr)
                continue
            hechas[registro["id"]] = registro
    return hechas


def medir(cliente: ClienteMaps, origen: str, destino: str) -> dict:
    datos = cliente.pedir("distancematrix", {
        "origins": origen, "destinations": destino,
        "mode": "driving", "units": "metric", "language": "es",
    })
    if datos.get("status") != "OK":
        return {"ok": False, "motivo": f"respuesta {datos.get('status')}"}

    filas = datos.get("rows") or []
    elementos = filas[0].get("elements") if filas else []
    if not elementos:
        return {"ok": False, "motivo": "respuesta sin elementos"}

    elemento = elementos[0]
    if elemento.get("status") != "OK":
        # ZERO_RESULTS aqui suele ser una ruta que no existe en auto, como
        # Margarita contra tierra firme.
        return {"ok": False, "motivo": f"elemento {elemento.get('status')}"}

    return {"ok": True,
            "km": round(elemento["distance"]["value"] / 1000, 1),
            "segundos": elemento["duration"]["value"]}


def procesar(cliente: ClienteMaps, pendientes: list[dict], tamano: int) -> dict[int, dict]:
    CHECKPOINT.parent.mkdir(parents=True, exist_ok=True)
    hechas = cargar_checkpoint()
    porhacer = [f for f in pendientes if f[COL_ID] not in hechas]
    if hechas:
        print(f"Reanudando: {len(hechas)} filas ya medidas")
    print(f"A medir en esta corrida: {len(porhacer)} de {len(pendientes)} candidatas\n")

    archivo = CHECKPOINT.open("a", encoding="utf-8")
    try:
        for contador, fila in enumerate(porhacer, start=1):
            identificador = fila[COL_ID]
            origen, destino = punto_de(fila, "Origen"), punto_de(fila, "Destino")

            if not origen or not destino:
                resultado = {"ok": False, "motivo": "fila sin coordenada ni place_id"}
            else:
                resultado = medir(cliente, origen, destino)

            registro = {"id": identificador, "tipo": tipo_de_calculo(fila), **resultado}
            archivo.write(json.dumps(registro, ensure_ascii=False) + "\n")

            if contador % tamano == 0 or contador == len(porhacer):
                archivo.flush()
                os.fsync(archivo.fileno())
                print(f"  [checkpoint] {contador}/{len(porhacer)} "
                      f"— llamadas reales={cliente.llamadas_reales}")
    except LimiteExcedido as error:
        print(f"\n[CUOTA AGOTADA] {error}", file=sys.stderr)
        print(f"Checkpoint en {CHECKPOINT}. Reanudar con el mismo comando.", file=sys.stderr)
    except KeyboardInterrupt:
        print("\n[interrumpido] guardando checkpoint...", file=sys.stderr)
    finally:
        archivo.flush()
        os.fsync(archivo.fileno())
        archivo.close()

    return cargar_checkpoint()


def escribir_columnas(libro: openpyxl.Workbook, encabezados: list[str],
                      filas: list[dict], medidas: dict[int, dict], ruta: Path) -> None:
    """Agrega las columnas al final de la hoja, sin tocar lo que ya estaba."""
    hoja = libro[hoja_de_datos(libro)]
    # Reusa las columnas si ya estan, para que volver a correr sobre el mismo
    # libro actualice los valores en vez de agregar un juego duplicado.
    inicio = (encabezados.index(COL_DISTANCIA) + 1 if COL_DISTANCIA in encabezados
              else len(encabezados) + 1)
    for desplazamiento, titulo in enumerate((COL_DISTANCIA, COL_DURACION, COL_TIPO)):
        celda = hoja.cell(row=1, column=inicio + desplazamiento, value=titulo)
        celda.font = hoja.cell(row=1, column=1).font.copy()
        celda.fill = hoja.cell(row=1, column=1).fill.copy()
        celda.alignment = hoja.cell(row=1, column=1).alignment.copy()
        hoja.column_dimensions[celda.column_letter].width = 22

    for fila in filas:
        medida = medidas.get(fila[COL_ID])
        numero = fila["__fila__"]
        if not medida:
            continue
        if medida.get("ok"):
            hoja.cell(row=numero, column=inicio, value=medida["km"])
            hoja.cell(row=numero, column=inicio + 1,
                      value=formatear_duracion(medida["segundos"]))
            # Un trayecto de cero puede ser real (la fuente repite el mismo texto
            # en ambos extremos, o el servicio es 'a disposicion' sin ruta fija)
            # o el sintoma de que dos lugares distintos colapsaron al mismo
            # punto. No se distinguen solos, asi que se marcan para revision.
            tipo = ("Revisar · ambos extremos resolvieron al mismo punto"
                    if medida["km"] == 0 else medida.get("tipo", ""))
        else:
            hoja.cell(row=numero, column=inicio, value="Sin ruta")
            hoja.cell(row=numero, column=inicio + 1, value=medida.get("motivo", "Sin ruta"))
            tipo = medida.get("tipo", "")
        hoja.cell(row=numero, column=inicio + 2, value=tipo)

    from openpyxl.cell.cell import TYPE_STRING
    for hoja_libro in libro.worksheets:
        for fila_libro in hoja_libro.iter_rows():
            for celda in fila_libro:
                if celda.data_type == "e" and celda.value is not None:
                    celda.data_type = TYPE_STRING
    libro.save(ruta)


def reportar(candidatas_filas: list[dict], medidas: dict[int, dict]) -> None:
    logradas = {i: m for i, m in medidas.items() if m.get("ok")}
    fallidas = {i: m for i, m in medidas.items() if not m.get("ok")}
    print("\n" + "=" * 70)
    print("RESULTADO")
    print("=" * 70)
    print(f"Candidatas (ambos extremos en Match) : {len(candidatas_filas)}")
    print(f"Con distancia y tiempo calculados    : {len(logradas)}")
    print(f"Fallidas                             : {len(fallidas)}")

    if fallidas:
        motivos: dict[str, int] = {}
        for medida in fallidas.values():
            motivos[medida.get("motivo", "?")] = motivos.get(medida.get("motivo", "?"), 0) + 1
        print("\nMotivos de fallo:")
        for motivo, cuenta in sorted(motivos.items(), key=lambda x: -x[1]):
            print(f"   {cuenta:>4}  {motivo}")

    tipos: dict[str, int] = {}
    for medida in logradas.values():
        tipos[medida.get("tipo", "?")] = tipos.get(medida.get("tipo", "?"), 0) + 1
    print("\nTipo de cálculo (sobre las logradas):")
    for tipo, cuenta in sorted(tipos.items(), key=lambda x: -x[1]):
        print(f"   {cuenta:>4}  {tipo}")


def main() -> None:
    analizador = argparse.ArgumentParser(description=__doc__)
    analizador.add_argument("--libro", type=Path, required=True)
    analizador.add_argument("--sublote", type=int, default=125,
                            help="Filas entre grabaciones de checkpoint")
    analizador.add_argument("--qps", type=float, default=15.0)
    analizador.add_argument("--solo-reporte", action="store_true",
                            help="No llama a la API: reporta el checkpoint existente")
    args = analizador.parse_args()

    libro, encabezados, filas = leer_hoja(args.libro)
    ya_tiene = [c for c in (COL_DISTANCIA, COL_DURACION, COL_TIPO) if c in encabezados]
    if ya_tiene and not args.solo_reporte:
        print(f"[aviso] la hoja ya trae {', '.join(ya_tiene)}; se actualizan en su sitio.")

    elegidas = candidatas(filas)
    print(f"Hoja {hoja_de_datos(libro)!r}: {len(filas)} filas | candidatas con ambos extremos en Match: "
          f"{len(elegidas)}")

    if args.solo_reporte:
        medidas = cargar_checkpoint()
    else:
        clave = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
        if not clave:
            raise SystemExit("Falta GOOGLE_MAPS_API_KEY en el entorno.")
        cliente = ClienteMaps(clave, CacheRespuestas(CACHE), qps=args.qps)
        medidas = procesar(cliente, elegidas, args.sublote)

    escribir_columnas(libro, encabezados, filas, medidas, args.libro)
    # El checkpoint es unico para todos los anios, asi que el reporte se acota a
    # las candidatas de este libro; si no, cada corrida sumaria las anteriores.
    del_libro = {i: m for i, m in medidas.items()
                 if i in {f[COL_ID] for f in elegidas}}
    reportar(elegidas, del_libro)
    print(f"\nLibro actualizado en sitio: {args.libro}")


if __name__ == "__main__":
    main()
