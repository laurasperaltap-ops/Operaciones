# Resumen de corrida — 2024, 2025 y 2026

Fuente: `Base rentas analisis_ DATA DIRECCIONES.xlsx`, hoja `Historico` (2695 filas).
Entregables: `salidas/Entregable_direcciones_<año>.xlsx`.
Reanudable desde: `checkpoints/<año>.jsonl` y `checkpoints/distancias.jsonl`.

## Cobertura por año

| | 2024 | 2025 | 2026 |
|---|---:|---:|---:|
| Filas | 1018 | 954 | 720 |
| `Ciudad` nula en la fuente | 677 (66.5 %) | 343 (36.0 %) | 0 |
| Origen con match | **1003 (98.5 %)** | **900 (94.3 %)** | **676 (93.9 %)** |
| Destino con match | **988 (97.1 %)** | **906 (95.0 %)** | **665 (92.4 %)** |
| Ambos extremos resueltos | 979 | 873 | 650 |
| Con distancia y tiempo | **975** | **869** | **650** |
| Sin ruta en auto | 4 | 4 | 0 |

Total: **2502 filas** con ambos extremos resueltos y **2494 con distancia y duración**.

2024 y 2025 salen mejor que 2026 pese a tener la columna `Ciudad` incompleta,
porque sus textos de `Origen`/`Destino` son mucho más completos: 0.2 % y 1.4 % de
nulos contra 3.3 % en 2026, y con descripciones más específicas.

## La falta de `Ciudad` no degradó el resultado

Era la preocupación principal antes de arrancar 2024. Medido sobre el primer
sublote, los lados resueltos **sin** ciudad declarada dieron 86.4 % de match
contra 83.3 % de los que **sí** la tenían. Una ciudad declarada incorrecta sesga
la consulta más de lo que un ancla correcta ayuda. Las filas resueltas sin ancla
quedan marcadas en la columna `Resuelto sin ancla de ciudad` para poder aislarlas.

## Dos reglas que se retiraron por medir mal

**Plus-code en `formatted_address`.** Se probó como síntoma de resultado vacío y
no lo es: Google omite el nombre del lugar de esa cadena incluso cuando resolvió
bien. `Estadio Monumental`, `Cocodrilos Sports Park` y `Palacio de las Academias`
resuelven correctamente y los tres aparecen como plus-code. La regla mandaba a
revisión 47 filas de 2026 y 66 del primer sublote de 2024 sin separar ni una
mala. Al quitarla, 2026 subió de 93.2 %/90.0 % a 93.9 %/92.4 %.

**Tipos faltantes.** `town_square` y `plus_code` no estaban en la lista de tipos
precisos, así que 12 plazas reales (`Plaza Altamira`, `Plaza Francia`,
`Plaza Los Palos Grandes`) y 2 filas que traen un plus-code como dirección caían
en un bucket `desconocida`. Ambos son ubicaciones precisas.

## Contrato de salida

Las columnas originales se conservan intactas y en orden; `ID cotización` es la
clave. Se agregan 14 columnas por lado (Origen y Destino), más `Conflicto ciudad
declarada` y las tres de distancia:

`Estado match` · `Precisión` · `Fuente` · `Estado` · `Municipio` · `Parroquia` ·
`Ciudad geocodificada` · `Dirección normalizada` · `Latitud` · `Longitud` ·
`Consulta usada` · `Place ID` · `Motivo sin match` · `Resuelto sin ancla de ciudad`,
y al final `Distancia (KM)` · `Duración estimada (HH:MM)` · `Distancia · Tipo de cálculo`.

Cuando no hay evidencia del geocoder, los campos llevan el literal `Sin match`;
nunca se completan por inferencia.

## Controles de sanidad sobre las 2494 distancias

- Ninguna velocidad promedio supera 110 km/h: no hay rutas físicamente imposibles.
- Las más largas son correctas: Guatire → Santa Elena de Uairén 1208 km / 17h43
  (frontera con Brasil), Táchira → La Guaira 853 km, Caracas → Maracaibo 709 km.
- Medianas coherentes con transporte urbano: 17.3 km (2024), 16.4 (2025), 11.6 (2026).

## Limitaciones conocidas

1. **Parroquia no está validada contra padrón oficial.** El campo toma
   `administrative_area_level_3`, `sublocality` o `neighborhood`, y a veces eso es
   una urbanización o un sector en vez de la parroquia civil. **Agregar por
   parroquia no es confiable todavía.** Requiere el padrón del INE/CNE.

2. **Trayectos de 0 km.** 36 filas entre los tres años. Parte son legítimas —la
   fuente repite el mismo texto en ambos extremos, o el servicio es "a
   disposición" sin ruta fija— y parte son dos lugares distintos que el geocoder
   colapsó a un punto. Van marcadas en `Distancia · Tipo de cálculo`.

3. **Rutas con tramo de ferry.** Google devuelve ruta en auto hacia Margarita
   (p. ej. Maracaibo → Nueva Esparta, 1171 km) incluyendo el cruce marítimo. La
   duración no contempla la espera del ferry.

4. **Topónimo distinto al declarado dentro del mismo estado.** "Distribuidora
   Ciudad Ojeda" con Ciudad = Maracaibo resuelve dentro de Maracaibo sin que se
   detecte. Requiere un listado de localidades venezolanas para tokenizar.

5. **Instalaciones recurrentes de clientes.** "Distribuidora Maracaibo Sur." y
   "Distr MCBO SUR" son probablemente la misma instalación y resuelven a puntos
   distintos. Una tabla de direcciones fijas lo eliminaría, pero requiere
   confirmación del negocio.

## Cómo reanudar

```bash
set -a && . ./.env && set +a
.venv/bin/python tools/geocodificar.py --anio <año> --qps 20
PYTHONPATH=tools .venv/bin/python tools/armar_entregable.py --anio <año>
.venv/bin/python tools/calcular_distancias.py --libro salidas/Entregable_direcciones_<año>.xlsx
```

Las filas ya resueltas se saltan por `ID cotización`. Para reprocesar con otra
heurística, borrar el `checkpoint` del año y volver a correr: el caché
`checkpoints/cache_geocoding.sqlite` evita pagar de nuevo las llamadas.
