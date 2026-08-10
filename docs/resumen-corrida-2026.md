# Resumen de corrida — año 2026

Fuente: `Base rentas analisis_ DATA DIRECCIONES.xlsx`, hoja `Historico`.
Entregable: `salidas/2026_direcciones_normalizadas.xlsx` (720 filas × 90 columnas).
Reanudable desde: `checkpoints/2026.jsonl`.

## Cobertura

| Resultado | Origen | Destino |
|---|---:|---:|
| Match exacto | 312 | 280 |
| Match aproximado | 306 | 176 |
| Match a nivel ciudad | 37 | 180 |
| **REVISAR** · coordenada sin nombre | 20 | 27 |
| **REVISAR** · fuera del área declarada | 10 | 12 |
| Sin match · no halló el lugar | 2 | 12 |
| Sin match · resolvió al país | 6 | 7 |
| Sin match · celda vacía en la fuente | 24 | 26 |
| Sin match · sin resultados | 3 | 0 |

- **Origen:** 655 con match (91.0 %), 30 a revisar, 35 sin match.
- **Destino:** 636 con match (88.3 %), 39 a revisar, 45 sin match.
- **Filas con ambos extremos resueltos: 606 / 720 (84.2 %).**
- Filas marcadas con conflicto de ciudad declarada: 57.

De los 80 "sin match", **50 son celdas vacías en la fuente**: no hay dirección que
buscar. Los sin-match atribuibles al proceso son 30.

## Estados cubiertos

15 estados. Miranda (530), Distrito Capital (478), La Guaira (159), Zulia (122),
Carabobo (26), Aragua (19), Lara (6), Anzoátegui (4), Falcón (4), y el resto con
volumen menor.

## Consumo

Sin llamadas nuevas en la corrida final: las 2034 resoluciones salieron del caché
SQLite. El total gastado contra la API durante todo el desarrollo, incluidas las
cinco reprocesadas por cambios de heurística, fue de aproximadamente **1350
llamadas** para 1440 direcciones.

## Contrato de salida

Las 63 columnas originales se conservan intactas y en orden; `ID cotización`
funciona como clave. Se agregan 13 columnas por lado (Origen y Destino) más una
de conflicto:

`Estado match` · `Precisión` · `Fuente` · `Estado` · `Municipio` · `Parroquia` ·
`Ciudad geocodificada` · `Dirección normalizada` · `Latitud` · `Longitud` ·
`Consulta usada` · `Place ID` · `Motivo sin match`, y `Conflicto ciudad declarada`.

Cuando no hay evidencia del geocoder, los campos llevan el literal `Sin match`;
nunca se completan por inferencia.

## Limitaciones conocidas

1. **Parroquia no está validada contra el padrón oficial.** El campo toma
   `administrative_area_level_3`, `sublocality` o `neighborhood` según lo que
   devuelva Google, y a veces eso es una urbanización o un sector ("La Campiña",
   "Terrazas del Ávila") en vez de la parroquia civil. Validarlo requiere el
   padrón del INE/CNE, que no está disponible en esta corrida. **Agregar por
   parroquia no es confiable todavía.**

2. **Topónimo distinto al declarado dentro del mismo estado.** Un texto como
   "Distribuidora Ciudad Ojeda" con Ciudad = Maracaibo puede resolver dentro de
   Maracaibo sin que se detecte, porque el resultado sí trae calle y sí cae en el
   área declarada. Detectarlo requiere un listado de localidades venezolanas
   contra el cual tokenizar el texto.

3. **Nombres sin normalizar en la salida.** Google devuelve "Caracas
   Metropolitan District" y "Municipio Autónomo Valencia"; se conservan tal cual
   vinieron. Se comparan normalizados internamente, pero un `join` externo contra
   esos valores necesita un diccionario de reemplazos.

4. **Enlaces de Google Maps en el texto libre.** Dos filas traen
   `https://maps.app.goo.gl/...` como dirección. Hoy quedan sin match; seguir el
   redirect daría la coordenada exacta sin gastar geocoding.

5. **Instalaciones recurrentes de clientes.** "Distribuidora Maracaibo Sur." y
   "Distr MCBO SUR" son probablemente la misma instalación y resuelven a puntos
   distintos. Una tabla de direcciones fijas de clientes frecuentes eliminaría
   ese ruido de raíz, pero requiere confirmación del negocio.

## Cómo reanudar

```bash
set -a && . ./.env && set +a
.venv/bin/python tools/geocodificar.py --anio 2026 --qps 20 \
    --exportar salidas/2026_direcciones_normalizadas.xlsx
```

Las filas ya resueltas se saltan por `ID cotización`. Para reprocesar con otra
heurística, borrar `checkpoints/2026.jsonl` y volver a correr: el caché
`checkpoints/cache_geocoding.sqlite` evita pagar de nuevo las llamadas.
