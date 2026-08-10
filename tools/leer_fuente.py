#!/usr/bin/env python3
"""Lectura y perfilado de la base de direcciones (fases 0 y 1).

Lee el archivo fuente en cualquiera de sus formatos (.numbers nativo de Apple,
.xlsx o .csv), lo normaliza a un DataFrame y reporta el perfilado que exige la
fase 1: total de filas, filas del ejercicio 2026 y nulos en Origen/Destino/Ciudad.

No aplica ninguna heuristica de matching ni toca el geocoder: su unico trabajo es
dejar la fuente legible y decir que hay dentro.

Uso:
    python tools/leer_fuente.py <ruta_fuente> [--exportar-a salidas/base.csv]
    python tools/leer_fuente.py <ruta_fuente> --hoja "Sheet 1" --tabla "Table 1"
"""

from __future__ import annotations

import argparse
import sys
import unicodedata
from pathlib import Path

import pandas as pd

# Columnas cuyo perfilado pide la fase 1. Se buscan por nombre normalizado
# (sin acentos, minusculas) para tolerar variantes de tipeo en la fuente.
COLUMNAS_PERFILADAS = ("origen", "destino", "ciudad")


def normalizar(texto: object) -> str:
    """Minusculas sin acentos ni espacios extremos, para comparar encabezados."""
    if texto is None:
        return ""
    sin_acentos = unicodedata.normalize("NFKD", str(texto))
    sin_acentos = "".join(c for c in sin_acentos if not unicodedata.combining(c))
    return sin_acentos.strip().lower()


def _tablas_numbers(ruta: Path):
    """Devuelve [(hoja, tabla, filas)] de un documento Apple Numbers."""
    from numbers_parser import Document

    documento = Document(str(ruta))
    for hoja in documento.sheets:
        for tabla in hoja.tables:
            yield hoja.name, tabla.name, list(tabla.rows(values_only=True))


def listar_tablas(ruta: Path) -> None:
    """Imprime el inventario de hojas/tablas: fase 0, antes de elegir cual leer."""
    print(f"Inventario de {ruta.name}:")
    for hoja, tabla, filas in _tablas_numbers(ruta):
        columnas = len(filas[0]) if filas else 0
        print(f"  hoja={hoja!r} tabla={tabla!r} filas={len(filas)} columnas={columnas}")


def _fila_encabezado(filas: list[list]) -> int:
    """Indice de la primera fila con >=2 celdas no vacias: el encabezado real.

    Las tablas de Numbers suelen traer filas de titulo o vacias arriba, asi que
    tomar la fila 0 a ciegas produce encabezados nulos.
    """
    for indice, fila in enumerate(filas):
        no_vacias = sum(1 for celda in fila if celda is not None and str(celda).strip())
        if no_vacias >= 2:
            return indice
    return 0


def leer_numbers(ruta: Path, hoja: str | None, tabla: str | None) -> pd.DataFrame:
    candidatas = [
        (h, t, f)
        for h, t, f in _tablas_numbers(ruta)
        if (hoja is None or h == hoja) and (tabla is None or t == tabla)
    ]
    if not candidatas:
        raise SystemExit(f"No hay hoja/tabla que coincida con hoja={hoja!r} tabla={tabla!r}")

    # Sin filtro explicito, la tabla con mas filas es la base de datos real.
    nombre_hoja, nombre_tabla, filas = max(candidatas, key=lambda c: len(c[2]))
    if len(candidatas) > 1:
        print(f"[aviso] {len(candidatas)} tablas; se usa la mayor: "
              f"hoja={nombre_hoja!r} tabla={nombre_tabla!r}", file=sys.stderr)
    if not filas:
        raise SystemExit(f"La tabla {nombre_tabla!r} esta vacia")

    inicio = _fila_encabezado(filas)
    encabezado = [str(c).strip() if c is not None else "" for c in filas[inicio]]
    marco = pd.DataFrame(filas[inicio + 1:], columns=encabezado)
    # Columnas sin nombre = relleno de la cuadricula de Numbers, no datos.
    marco = marco.loc[:, [c for c in marco.columns if c != ""]]
    return marco.dropna(how="all")


def cargar(ruta: Path, hoja: str | None = None, tabla: str | None = None) -> pd.DataFrame:
    sufijo = ruta.suffix.lower()
    if sufijo == ".numbers":
        return leer_numbers(ruta, hoja, tabla)
    if sufijo in (".xlsx", ".xlsm"):
        return pd.read_excel(ruta, sheet_name=hoja or 0)
    if sufijo == ".csv":
        return pd.read_csv(ruta)
    raise SystemExit(f"Formato no soportado: {sufijo}")


def ubicar(marco: pd.DataFrame, clave: str) -> str | None:
    """Busca una columna cuyo nombre normalizado contenga `clave`."""
    for columna in marco.columns:
        if clave in normalizar(columna):
            return columna
    return None


def contar_2026(marco: pd.DataFrame) -> tuple[int, str | None]:
    """Filas del ejercicio 2026 y la columna de la que se dedujo el año.

    Prefiere una columna de año explicita; si no existe, cae a la primera
    columna de fecha parseable. Devuelve (-1, None) si no hay de donde deducirlo,
    para no reportar un cero que parezca un dato real.
    """
    for columna in marco.columns:
        nombre = normalizar(columna)
        if nombre in ("anio", "ano", "year", "ejercicio") or "anio" in nombre or "año" in str(columna).lower():
            valores = pd.to_numeric(marco[columna], errors="coerce")
            return int((valores == 2026).sum()), columna

    for columna in marco.columns:
        if "fecha" in normalizar(columna) or "date" in normalizar(columna):
            fechas = pd.to_datetime(marco[columna], errors="coerce")
            if fechas.notna().any():
                return int((fechas.dt.year == 2026).sum()), columna

    return -1, None


def perfilar(marco: pd.DataFrame, ruta: Path) -> None:
    print("=" * 64)
    print(f"FASE 0-1 · PERFILADO · {ruta.name}")
    print("=" * 64)
    print(f"Filas totales : {len(marco)}")
    print(f"Columnas      : {len(marco.columns)}")
    print()
    print("Columnas detectadas:")
    for columna in marco.columns:
        nulos = int(marco[columna].isna().sum()) + int((marco[columna].astype(str).str.strip() == "").sum())
        print(f"  - {columna!r:<40} nulos/vacios={nulos}")
    print()

    filas_2026, origen_anio = contar_2026(marco)
    if filas_2026 < 0:
        print("Filas 2026    : NO DETERMINADO (sin columna de anio ni de fecha parseable)")
    else:
        print(f"Filas 2026    : {filas_2026}   (deducido de {origen_anio!r})")
    print()

    print("Nulos en columnas criticas (fase 1):")
    for clave in COLUMNAS_PERFILADAS:
        columna = ubicar(marco, clave)
        if columna is None:
            print(f"  - {clave:<10}: COLUMNA NO ENCONTRADA en la fuente")
            continue
        serie = marco[columna]
        vacios = int(serie.isna().sum()) + int((serie.astype(str).str.strip().isin(["", "nan", "None"])).sum())
        pct = (vacios / len(marco) * 100) if len(marco) else 0.0
        print(f"  - {clave:<10}: {vacios} nulos/vacios ({pct:.1f}%)  [columna {columna!r}]")
    print("=" * 64)


def main() -> None:
    analizador = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    analizador.add_argument("fuente", type=Path, help="Ruta al .numbers / .xlsx / .csv")
    analizador.add_argument("--hoja", default=None, help="Nombre de hoja a leer")
    analizador.add_argument("--tabla", default=None, help="Nombre de tabla a leer (solo .numbers)")
    analizador.add_argument("--listar", action="store_true", help="Solo inventariar hojas/tablas y salir")
    analizador.add_argument("--exportar-a", type=Path, default=None, help="Vuelca la fuente a .csv o .xlsx")
    args = analizador.parse_args()

    if not args.fuente.exists():
        raise SystemExit(f"No existe el archivo: {args.fuente}")

    if args.listar:
        listar_tablas(args.fuente)
        return

    marco = cargar(args.fuente, args.hoja, args.tabla)
    perfilar(marco, args.fuente)

    if args.exportar_a:
        args.exportar_a.parent.mkdir(parents=True, exist_ok=True)
        if args.exportar_a.suffix.lower() == ".csv":
            marco.to_csv(args.exportar_a, index=False)
        else:
            marco.to_excel(args.exportar_a, index=False)
        print(f"\nExportado a {args.exportar_a}")


if __name__ == "__main__":
    main()
