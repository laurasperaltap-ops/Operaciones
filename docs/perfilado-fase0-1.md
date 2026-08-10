# Perfilado fases 0-1 — Base rentas análisis · DATA DIRECCIONES

Fuente: `Base rentas analisis_ DATA DIRECCIONES.xlsx` (export de Numbers, hoja única `Historico`).
Corrida: `python tools/leer_fuente.py datos/Base_rentas_analisis_DATA_DIRECCIONES.xlsx`

## Fase 0 — verificación

| Chequeo | Resultado |
|---|---|
| Hojas | 1 (`Historico`) |
| Rango bruto | 2862 × 63 (incluye filas de relleno al final) |
| Filas con datos | **2695** |
| Columnas | **63** |
| Geocoding API | `OK` — devuelve `administrative_area_level_1/2`, `locality`, `sublocality`, `country` |
| Distance Matrix API | `OK` |

## Fase 1 — perfilado

### Distribución por año (columna `Año`)

| Año | Filas |
|---|---|
| 2024 | 1018 |
| 2025 | 954 |
| **2026** | **720** |
| sin año | 3 |

### Nulos en columnas críticas

| Columna | Total (2695) | Solo 2026 (720) |
|---|---|---|
| `Origen ` *(con espacio final en el nombre)* | 40 (1.5 %) | 24 (3.3 %) |
| `Destino` | 41 (1.5 %) | 26 (3.6 %) |
| `Ciudad` | 1023 (38.0 %) | **0 (0 %)** |

Los 1023 nulos de `Ciudad` están **todos** en 2024-2025. El lote 2026 tiene ciudad completa
en las 720 filas, y 693 de ellas traen `Origen` y `Destino` presentes a la vez.

### Clave primaria

`ID cotización`: `int64`, sin nulos, **2695 valores únicos**. Sirve como clave de checkpoint
y de reanudación sin ambigüedad.

## Hallazgos que condicionan la heurística

1. **`Ciudad` no está normalizada.** Conviven `Caracas` (1398), `CARACAS` (59), `Caracas `
   con espacio final (6) y `caracas` (3); igual con `Valencia`/`VALENCIA`. Hay además una
   categoría atrapatodo `Otra (Venezuela)` (7 filas).

2. **`Origen`/`Destino` son texto libre tipo POI, no direcciones postales.** Ejemplos reales:
   `Aeropuerto de Maiquetia`, `Plaza Altamira`, `Cubo negro - Bangente`,
   `sabaneta calle 99c cerca de la clínica zulia`. La calidad es dispar: nombres de hitos
   reconocibles conviven con descripciones coloquiales.

3. **`Ciudad` contradice al par Origen/Destino en algunas filas.** Casos detectados en 2026:
   - `ID 1972`: `Ciudad = Caracas`, pero origen `Av. Bolívar Norte, urb. Ca…` y destino
     `Parque temático Isla La Cumaca` — ambos en Carabobo, no en Caracas.
   - `ID 1976`: `Ciudad = Caracas`, destino `Salón de Asambleas, Cúa, MIR` — Cúa es Miranda.

   Usar `Ciudad` como sesgo duro del geocoder arrastraría estos errores. Cómo se resuelve el
   conflicto es decisión de la sección 5 del documento del proyecto.
