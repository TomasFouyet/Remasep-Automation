# tests/fixtures/

Archivos binarios pequeños, sintéticos y sin PII usados por tests opt-in que
no pueden generarse en cada corrida (por ejemplo, porque requieren Excel real
en Windows). A diferencia de `data/local/` (plantillas/exports reales,
gitignored), estos archivos **sí se commitean**.

- `synthetic_workbook_open.xlsm` (F06, `tests/test_excel_com_integration.py`):
  workbook benigno con un único `Workbook_Open` que escribe la marca
  `"MARKER_FIRED"` en `A1`. Generado una vez con
  `scripts/build_workbook_open_fixture.py` (Windows + Excel). No existe
  todavía en este checkout — ver `NEEDS_WINDOWS_VALIDATION` en
  `docs/EXCEL_WRITER.md`.
