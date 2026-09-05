# Sprint 1.4 — Official Template Inventory & Alignment

## Objetivo

Inventariar la **plantilla oficial** `REMASEP 2026_V1.4.xlsm` (estructura, hojas,
fórmulas, protección, VBA) y producir *candidatos* de alineación entre el
generador legacy y la plantilla — sin confirmar ningún mapping ni ejecutar
macros.

## Problema que resolvía

El REMASEP final se entrega en el formato oficial MINSAL. Hay que conocer ese
formato (y en qué se parece o no al generador legacy) antes de diseñar el mapping
semántico.

## Implementación

`scripts/inventory_template.py` — `openpyxl` con `keep_vba=True`, sin COM.
Extrae estructura por hoja, grafo interno, celdas candidatas a ingreso
(desbloqueadas / con validación en hojas protegidas) y candidatos de alineación
por **contenido de etiqueta** (`exact_label` / `normalized_label` / `contextual`
/ `ambiguous`). Genera [`docs/TEMPLATE_ALIGNMENT.md`](../TEMPLATE_ALIGNMENT.md),
`template_fingerprint.json` y CSV.

## Archivos principales

- [`scripts/inventory_template.py`](../../scripts/inventory_template.py)
- [`docs/TEMPLATE_ALIGNMENT.md`](../TEMPLATE_ALIGNMENT.md)

## Decisiones técnicas

- VBA: sólo se reporta **presencia + tamaño + hash**; nunca se interpreta ni
  ejecuta.
- Alineación por **etiqueta**, no por desplazamiento de coordenadas; nunca se
  elige un candidato automáticamente.
- `input_candidates` = candidatos **estructurales**, no campos de ingreso
  confirmados.
- `template_fingerprint.json` (experimental) para detectar si una futura "V1.4"
  cambia de estructura sin cambiar de nombre.

## Tests

Workbooks sintéticos con hojas protegidas, validaciones de datos, celdas
desbloqueadas y un valor secreto que **no** debe aparecer en la salida.

## Resultado sobre workbook de referencia

`REMASEP 2026_V1.4.xlsm` (SHA256 `a59a335c…d60e64`):

| Métrica | Valor |
| --- | ---: |
| Hojas | 11 |
| Fórmulas | 11 466 |
| VBA (`xl/vbaProject.bin`) | presente, 49 664 bytes |
| Hojas sin protección | `CONTROL` (sólo) |
| Candidatos `exact_label` | 230 |
| Candidatos `ambiguous` | 1189 |
| Input candidates (estructurales) | 26 336 |

Hojas: `NOMBRE`, `REMASEP 01`, `URGENCIAS`, `REMASEP B1`, `B2 ANEXO`,
`REMASEP_OD`, `EyP_ET`, `TV_MI`, `SERV_SANGRE`, `CONTROL`, `MACROS`.

## Hallazgos

- `CONTROL` es el **verificador final** del formato oficial (única hoja sin
  proteger; todas las demás la alimentan con checks).
- `NOMBRE` alimenta el encabezado de todas las hojas de formulario;
  `B2 ANEXO` → `REMASEP B1` (83) y → `REMASEP 01` (8).
- El generador legacy y la plantilla comparten nombres de hoja
  (`REMASEP 01`, `B2 ANEXO`, `REMASEP_OD`) pero la plantilla tiene **8 hojas
  más** y muchas más celdas de ingreso.

## Limitaciones

- **El mapping oficial NO está implementado.** Sólo hay candidatos de etiqueta;
  1189 son ambiguos y requieren decisión humana.
- Los `input_candidates` se basan en protección/estilo/validación, no en
  confirmación del proceso real.

## Estado final

**COMPLETE.** Estructura oficial documentada; el mapping semántico queda para
después del inventario completo de fuentes (pre-Sprint 3) y Sprint 3.
