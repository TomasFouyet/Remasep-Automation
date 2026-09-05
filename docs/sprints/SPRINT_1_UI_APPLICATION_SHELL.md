# Sprint 1 UI — REMASEP Application Shell

## Objetivo

Maqueta de escritorio **navegable** en PySide6 que recorra el flujo completo del
proceso (cargar → validar → revisar → resultado) con datos ficticios, sin tocar
archivos reales ni COM.

## Problema que resolvía

Necesitábamos una superficie concreta para acordar el flujo de usuario y los
modelos de dominio de la UI antes de conectar el motor real.

## Implementación

- `src/remasep/ui/main_window.py` con `QStackedWidget` y 5 pantallas:
  Home · Nuevo reporte · Análisis · Revisión (excepciones) · Resultado.
- `src/remasep/services/mock_remasep.py` — `MockRemasepService` con datos
  ficticios; **modo demostración** para recorrer todo sin archivos.
- Componentes reutilizables: `file_selector`, `status_card`, `step_indicator`.
- Stylesheet central en `ui/styles.py`.
- Modelos de dominio de UI: `AnalysisResult`, `ValidationResult`,
  `ExceptionItem`, con invariantes
  (`total = classified + ignored + unclassified`).
- La pantalla de Resultado distingue **el resultado del análisis** del **estado
  posterior a la revisión** (proyección `ReviewOutcome` a partir de las
  decisiones de la sesión).

## Archivos principales

- [`src/remasep/ui/main_window.py`](../../src/remasep/ui/main_window.py)
- [`src/remasep/ui/screens/`](../../src/remasep/ui/screens/) — `home`,
  `new_report`, `analysis`, `exceptions`, `summary`
- [`src/remasep/ui/components/`](../../src/remasep/ui/components/)
- [`src/remasep/services/mock_remasep.py`](../../src/remasep/services/mock_remasep.py)

## Decisiones técnicas

- El botón **"Generar REMASEP" está deshabilitado** (requiere motor +
  integración Excel).
- El selector de archivos abre `QFileDialog` y registra nombre/tamaño, pero **no
  procesa** nada.
- Sin persistencia de reglas: las excepciones se resuelven sólo durante la
  sesión.
- Tests con Qt en modo **offscreen** (`tests/conftest.py`), sin display físico.

## Tests

Pantallas y componentes con `QT_QPA_PLATFORM=offscreen`; invariantes de los
modelos de UI; ordenamiento de señales del período por defecto.

## Resultado sobre workbook de referencia

N/A — la UI de este sprint sólo usa datos mock.

## Hallazgos

- **Grey lines**: una regla global `QWidget { background }` pintaba franjas
  grises detrás de cada `QLabel`. Corregido: fondo sólo en
  `QMainWindow`/`QStackedWidget`; `QLabel { background: transparent }`; nada de
  reglas globales de `QLabel` con `background-color`.
- Bugs semánticos de modelos corregidos en review: distinción
  `total/classified/ignored/unclassified`; período por defecto = mes anterior en
  hora local; proyección `ReviewOutcome` calculada, no fija.

## Limitaciones

- Todo mock; el modo **REAL** se conecta en Sprint 2.1.
- Sin generación oficial.

## Estado final

**COMPLETE.** Flujo `HOME → NEW REPORT → ANALYSIS → EXCEPTIONS → SUMMARY`
navegable; base para conectar el análisis Medinet real.
