# Ampliación de destinos Medinet 2026 (rama experimental)

Fuente productiva única: export directo `detalle_citas`; se aplica el mismo
período, agenda ambulatoria y filtro ESTADO que al manifiesto original. La
plantilla oficial en blanco conserva sus fórmulas; los nuevos destinos son
entradas literales, desbloqueadas y sin fórmula.

`scripts/extend_medinet_scope_2026.py` genera de forma idempotente 85 filas
adicionales en `metric_catalog.csv` y `write_manifest.csv`, y actualiza sus
hashes en `bundle.yaml`. El contrato pasa de 1.122 a **1.207** escrituras. Las
1.122 originales se mantienen con los mismos valores: para septiembre 2026
su digest previo sigue siendo
`557ef5fb3171f2d9fe6c6e7832e4cd27540e542e8e051b80160fc8e00333117c`.

| Destinos nuevos | Cantidad | Regla Medinet |
|---|---:|---|
| `REMASEP 01!D15:M16` | 20 | Consultas médicas de especialidad y atenciones de Enfermería por edad al día de cita y sexo. Sólo las diez celdas de cada una de las filas 15 y 16. |
| `REMASEP 01!D55:V55`, `D75:V75`, `D89:V89` | 57 | Genética, Cirugía Plástica y Otorrinolaringología por bandas quinquenales y sexo; la fuente es TIPO_DE_CITA. |
| `B2 ANEXO!C1251`, `C1255`, `C1256`, `C1258`, `C1267`, `C1303`, `C1304`, `C1307` | 8 | Código de PRESTACION para seis procedimientos; TIPO_DE_CITA y SUCURSAL para las dos rehabilitaciones. |

En el export de septiembre (SHA256
`18b11ce185f7231c74d0d3345095b88196776ffa03efe8d47096692a696066e9`):
1.226 registros elegibles, 1.207 escrituras, `completeness.ok=true`.
Médico 108 (52 hombres, 56 mujeres) y Enfermería 58 (32 hombres, 26 mujeres).
Los ocho procedimientos de B2 son `14,1,6,15,1,4,135,55`, todos iguales a
los valores históricos. Para rehabilitación, Santiago `81+54=135` y
Telemedicina `20+35=55`, separadas por SUCURSAL.

**Diferencias de edad:** de las 77 celdas nuevas en REMASEP 01, 48 coinciden
exactamente con el histórico y 29 difieren. Menores de diez años: Médico 55
frente a 81 histórico; Enfermería 22 frente a 47. La edad se recalcula desde
FECHA_NACIMIENTO a DIA_CITA, igual que el productor existente. No se copian
las edades cacheadas del generador legacy.

**Fuera de alcance de esta ampliación:** `REMASEP_OD!Q61` y `S81` (cirugías);
los 38 campos complementarios pendientes en CONTROL y los metadatos de la
portada. No se han demostrado sus reglas con el export ambulatorio. El libro
completo continúa siendo un borrador, **no apto para envío**. La validación
del writer real, CONTROL, VBA y output guardado requiere Windows y Excel COM.

Pruebas locales: `PYTHONPATH=src python -m pytest -q -m 'not excel'`:
980 aprobadas, 12 omitidas, 4 pruebas Excel excluidas. Prueba sintética nueva
verifica edades, sexo, sucursal, ESTADO y seis destinos representativos sin
depender de datos privados. Se comprobó además que los 85 destinos están
desbloqueados y no contienen fórmulas en la plantilla oficial.
