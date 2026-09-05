# Sprint 1.1 — Workbook Inventory

## Objetivo

Inventariar de forma estática y **solo lectura** el workbook generador legacy
`GENERACION DATOS REMASEP.xlsx`: hojas, celdas, fórmulas y patrones de fórmula,
sin ejecutar Excel ni exportar valores de pacientes.

## Problema que resolvía

Antes de reproducir cualquier lógica hay que saber qué contiene realmente el
archivo con el que hoy se produce el REMASEP: cuántas hojas, cuántas fórmulas,
qué familias de fórmula y si usa constructs difíciles (named ranges, `INDIRECT`…).

## Implementación

`scripts/inventory_workbook.py` — `openpyxl` con `data_only=False` (conserva las
fórmulas), sin COM. Calcula SHA256 del archivo, recorre las hojas, normaliza cada
fórmula a un "patrón" y agrega conteos. Escribe CSV + `README.md` a
`artifacts/workbook_inventory/` (gitignored).

Correcciones posteriores del extractor de nombres de función: ignorar el
contenido dentro de literales de texto (`"EMBARAZADAS"` no es una función).

## Archivos principales

- [`scripts/inventory_workbook.py`](../../scripts/inventory_workbook.py)
- Artefactos: `artifacts/workbook_inventory/*.csv` (gitignored)

## Decisiones técnicas

- Solo estructura y texto de fórmulas; **cero** valores de celda en la salida.
- `data_only=False` para no depender de la cache de Excel en esta fase.
- Patrón de fórmula = fórmula con literales y coordenadas normalizados, para
  contar familias en lugar de fórmulas individuales.

## Tests

Tests con workbooks sintéticos (sin `data/local/`) que cubren extracción de
hojas/fórmulas/patrones y el filtrado de falsos nombres de función.

## Resultado sobre workbook de referencia

`GENERACION DATOS REMASEP.xlsx` (SHA256 `fc2e1536…d15b79`):

| Métrica | Valor |
| --- | ---: |
| Hojas | 4 |
| Celdas no vacías | 54 682 |
| Fórmulas | 22 642 |
| Patrones de fórmula | 39 |
| Named ranges | 0 |

Hojas: `REMASEP 01`, `B2 ANEXO`, `REMASEP_OD`, `Atenciones - Detalles de citas`.

## Hallazgos

- El workbook es 100 % fórmulas + datos crudos; sin named ranges ni constructs
  dinámicos.
- La hoja `Atenciones - Detalles de citas` tiene **2006 filas físicas**.
  **Sprint 2.1 comprobó después que esas 2006 filas NO son 2006 atenciones**:
  642 son filas estructuralmente vacías con las fórmulas `AC:AL` arrastradas más
  allá de las atenciones reales (1364). Ver
  [`docs/MEDINET_ANALYSIS.md`](../MEDINET_ANALYSIS.md).

## Limitaciones

- Inventario estructural, no semántico: no dice qué *significa* cada fórmula.
- Conteos por **fila física**, no por atención.

## Estado final

**COMPLETE.** Base de datos de fórmulas y patrones lista para el análisis de
dependencias (Sprint 1.2).
