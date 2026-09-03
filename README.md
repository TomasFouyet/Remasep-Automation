# REMASEP Automation

Aplicación desktop **local** para automatizar la preparación mensual del REMASEP de
Fundación Gantz a partir de fuentes institucionales: **Medinet**, **egresos
hospitalarios** y **datos de recursos / pabellones**.

El objetivo final es reemplazar el proceso manual actual —que hoy toma varias horas—
por un flujo guiado: **cargar fuentes → validar → revisar excepciones → generar el
workbook oficial**.

Este repositorio todavía **no** contiene reglas clínicas/REMASEP validadas ni un
mapping confirmado hacia las celdas del formulario oficial. Lo que existe hoy es el
resultado del *discovery* técnico de Sprint 1 más una maqueta de UI navegable.

---

## 1. Estado actual del proyecto

**Estado:** *Discovery técnico completado + UI shell funcional.*

- [x] Arquitectura base
- [x] Ingeniería inversa del workbook actual (Sprint 1.1)
- [x] Dependencias de fórmulas (Sprint 1.2)
- [x] Inventario de la lógica actual (Sprint 1.3)
- [x] Inventario de la plantilla oficial `REMASEP 2026 V1.4` (Sprint 1.4)
- [x] Alineación preliminar generador ↔ plantilla oficial (Sprint 1.4)
- [x] UI navegable con datos mock (UI Sprint 1)
- [ ] Análisis real de Medinet
- [ ] Equivalencia Python ↔ workbook legacy
- [ ] Mapping semántico hacia la plantilla oficial
- [ ] Integración Excel COM
- [ ] Integración de egresos
- [ ] Integración de recursos / pabellones
- [ ] Generación oficial
- [ ] CONTROL
- [ ] Empaquetado `.exe`

---

## 2. Hallazgos principales de Sprint 1

> Resumen. El detalle completo (CSV + diagramas) vive en `docs/` y en
> `artifacts/` tras ejecutar los scripts de inventario.

### Workbook generador — `data/local/GENERACION DATOS REMASEP.xlsx`

- 4 hojas: `REMASEP 01`, `B2 ANEXO`, `REMASEP_OD`, `Atenciones - Detalles de citas`.
- 22.642 fórmulas · 39 patrones de fórmula · 0 named ranges.

### Hoja de detalle (`Atenciones - Detalles de citas`)

- 2.006 atenciones analizadas en el workbook de referencia.
- 10 columnas derivadas `AC:AL`, con 1 patrón de fórmula por columna
  (todas las filas replican la misma fórmula).
- 27 reglas candidatas extraídas de las columnas de clasificación
  (`AG` 3, `AH` 4, `AI` 6, `AJ` 6, `AK` 3, `AL` 5).

**Columnas raw usadas transitivamente por las hojas output (7):**
`DIA CITA`, `FECHA NACIMIENTO`, `SEXO`, `SUCURSAL`, `ESPECIALIDAD`,
`TIPO DE CITA`, `PRESTACIÓN`.

**Transformaciones derivadas observadas:**

| Col | Qué hace (técnico) |
| --- | --- |
| `AC` | `TIPO DE CITA` + `SUCURSAL` |
| `AD` | `TIPO DE CITA` + `SEXO` |
| `AE` | `PRESTACIÓN` + `ESPECIALIDAD` + `SEXO` |
| `AF` | edad (a partir de fecha de cita y fecha de nacimiento) |
| `AG` | consultas médicas |
| `AH` | controles odontológicos de especialidad |
| `AI` | evaluaciones odontológicas de especialidad |
| `AJ` | controles ortodoncia |
| `AK` | controles ortopedia |
| `AL` | instalaciones ortopedia |

> ⚠️ Estas reglas **reproducen el comportamiento del workbook actual** y están
> **pendientes de validación funcional**. No son "reglas oficiales MINSAL".

**Observaciones que requieren confirmación del responsable funcional:**

- `ESTADO` existe pero no participa en ninguna dependencia hacia REMASEP.
- `MODALIDAD` existe pero no participa; el flag presencial/telemedicina parece
  derivarse de `TIPO DE CITA` + `SUCURSAL` (columna `AC`).
- `PRESTACIÓN REALIZADA` existe pero no participa; las reglas usan `PRESTACIÓN`.

Detalle: [`docs/CURRENT_LOGIC.md`](docs/CURRENT_LOGIC.md),
[`docs/CURRENT_EXCEL_FLOW.md`](docs/CURRENT_EXCEL_FLOW.md).

---

## 3. Plantilla oficial `REMASEP 2026 V1.4`

- 11 hojas detectadas: `NOMBRE`, `REMASEP 01`, `URGENCIAS`, `REMASEP B1`,
  `B2 ANEXO`, `REMASEP_OD`, `EyP_ET`, `TV_MI`, `SERV_SANGRE`, `CONTROL`, `MACROS`.
- Contiene **VBA** (macros; no se ejecutan ni se interpretan en este proyecto).
- 11.466 fórmulas.
- `CONTROL` es el **verificador final** (única hoja sin protección).
- Flujo interno observado: `B2 ANEXO` alimenta `REMASEP B1` y también `REMASEP 01`;
  `NOMBRE` alimenta los encabezados de las hojas de formulario; `CONTROL` recibe
  dependencias de varias hojas.

**Alineación generador ↔ plantilla** (por contenido de etiqueta, **no** por
desplazamiento de coordenadas):

- 230 candidatos `exact_label`.
- 1.189 candidatos `ambiguous` (nunca se elige uno de forma automática).
- El mapping definitivo será **semántico / estructural**, no posicional.

> Los ~26.336 *input candidates* detectados son **candidatos estructurales**
> (celdas desbloqueadas / con validación en hojas protegidas), **no** campos de
> ingreso confirmados. Requieren validación humana.

Detalle: [`docs/TEMPLATE_ALIGNMENT.md`](docs/TEMPLATE_ALIGNMENT.md).

---

## 4. Arquitectura

```text
EXTRACT
    Medinet / Egresos / Recursos
        ↓
TRANSFORM
    normalizar / validar / clasificar
        ↓
METRICS
    tabla larga semántica
        ↓
MAP
    métrica → destino según versión REMASEP
        ↓
LOAD
    Excel oficial / recálculo / CONTROL / guardar
```

Principios (ver [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)):

- La lógica de negocio **no** conoce coordenadas Excel (`G54`); el mapping por
  versión traduce cada métrica semántica a una celda/rango.
- Una prestación desconocida o una clasificación ambigua **bloquean** el
  procesamiento (0 matches → bloquear · 1 → continuar · >1 → bloquear).
- La plantilla MINSAL es el formato oficial de salida.

**Plataformas:**

| Entorno | Uso |
| --- | --- |
| Ubuntu / Linux | desarrollo, reglas, tests, análisis, UI |
| Windows + Microsoft Excel | integración COM y generación oficial final |

No se usa LibreOffice como sustituto de Microsoft Excel para la validación final.

---

## 5. UI actual (UI Sprint 1)

Maqueta navegable en **PySide6**, conectada a `MockRemasepService` (datos ficticios).

**Pantallas:** Home · Nuevo reporte · Análisis · Revisión · Resultado.

```text
HOME → Archivos → Análisis → Revisión → Resultado
```

Hoy:

- Funciona en Ubuntu con PySide6; incluye un **modo demostración** para recorrer
  todo el flujo sin archivos reales.
- El selector de archivos registra nombre/tamaño pero **no procesa** Medinet real.
- Las excepciones se pueden **resolver visualmente** durante la sesión; no hay
  persistencia de reglas.
- La pantalla de Resultado distingue **el resultado original del análisis** del
  **estado posterior a la revisión** (proyección `ReviewOutcome` calculada a
  partir de las decisiones de la sesión).
- El botón **"Generar REMASEP" está deshabilitado** (requiere el motor de
  procesamiento y la integración con Excel).

---

## 6. Cómo ejecutar

**Ubuntu:**

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e ".[dev]"
python -m remasep.main
```

**Tests y lint:**

```bash
pytest
ruff check .
```

Los tests de UI configuran Qt en modo **offscreen** automáticamente
(`tests/conftest.py`), por lo que no requieren display físico.

---

## 7. Estructura del repositorio

```text
src/remasep/
  adapters/        Medinet (preliminar), wrapper Excel COM (Windows)
  core/            normalización, motor de reglas, agregación (tabla larga)
  domain/          modelos
  services/        mock_remasep.py  (datos ficticios para la UI)
  ui/
    screens/       home · new_report · analysis · exceptions · summary
    components/     file_selector · status_card · step_indicator
    main_window.py · styles.py

scripts/           herramientas de ingeniería inversa (solo lectura)
  inventory_workbook.py        Sprint 1.1
  formula_refs.py              analizador de referencias A1
  analyze_dependencies.py      Sprint 1.2
  inventory_current_logic.py   Sprint 1.3
  inventory_template.py        Sprint 1.4

docs/              ARCHITECTURE · CURRENT_EXCEL_FLOW · CURRENT_LOGIC ·
                   TEMPLATE_ALIGNMENT · ROADMAP · CLIENT_PENDING
config/            plantillas de configuración versionada por REMASEP
tests/
```

Los artefactos generados por los scripts se escriben en `artifacts/` (gitignored).

---

## 8. Seguridad y privacidad

- **Procesamiento local.** Sin telemetría por defecto.
- Los archivos reales viven en `data/` y están **gitignored**.
- Las herramientas de ingeniería inversa se diseñaron para **no exportar valores
  de pacientes**: solo fórmulas, metadatos, encabezados y estructura.
- Los datos nominales no se incluyen en los inventarios.
- Los logs futuros no deben contener RUT, nombre ni fecha de nacimiento
  individuales.

---

## 9. Siguiente paso — Sprint 2.1: Medinet Real Analysis MVP

Seleccionar un archivo Medinet **real** desde la UI y mostrar:

- total de registros;
- validación de estructura;
- validación de período;
- cálculo de edad;
- clasificaciones legacy actuales (las 27 reglas candidatas);
- warnings;
- estadísticas reales del archivo.

Todavía **no** generará el REMASEP oficial.

Luego: equivalencia con el Excel legacy → mapping semántico → integración Excel
COM → egresos / recursos → CONTROL.

---

## 10. Documentación relacionada

| Documento | Para qué sirve |
| --- | --- |
| [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) | Componentes y flujo Extract→Transform→Metrics→Map→Load |
| [`docs/CURRENT_EXCEL_FLOW.md`](docs/CURRENT_EXCEL_FLOW.md) | Grafo de dependencias entre hojas del workbook generador (observado) |
| [`docs/CURRENT_LOGIC.md`](docs/CURRENT_LOGIC.md) | Lógica actual de la hoja de detalle: columnas, transformaciones, reglas candidatas |
| [`docs/TEMPLATE_ALIGNMENT.md`](docs/TEMPLATE_ALIGNMENT.md) | Estructura de la plantilla oficial y candidatos de alineación |
| [`docs/ROADMAP.md`](docs/ROADMAP.md) | Plan por sprints |
| [`docs/CLIENT_PENDING.md`](docs/CLIENT_PENDING.md) | Información pendiente de confirmar con Fundación Gantz / MINSAL |
