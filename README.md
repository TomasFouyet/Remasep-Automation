# REMASEP Automation

Aplicación local para automatizar la generación mensual de REMASEP a partir de fuentes institucionales
como Medinet, egresos hospitalarios y datos de pabellones/recursos.

## Estado

**Sprint 0 / Base técnica**

Este repositorio todavía NO contiene reglas clínicas/REMASEP definitivas ni mappings reales hacia celdas
del formulario oficial. Esos elementos se incorporarán solamente después de validarlos con los archivos
históricos y el proceso actual.

## Principios

- Procesamiento local.
- La lógica de negocio no depende de coordenadas Excel.
- Una prestación desconocida bloquea el procesamiento.
- Una clasificación ambigua bloquea el procesamiento.
- La plantilla MINSAL se trata como formato oficial de salida.
- Para el `.xlsm` final se utilizará Microsoft Excel vía COM en Windows.
- No se deben guardar datos nominales de pacientes en logs.
- Cada versión de REMASEP tendrá su propia configuración/mapping.

## Flujo objetivo

```text
EXTRACT
Medinet / Egresos / Recursos
        ↓
TRANSFORM
normalizar → validar → clasificar → agregar
        ↓
METRICS
tabla larga de métricas semánticas
        ↓
MAP
métrica → hoja/celda según versión
        ↓
LOAD
Excel oficial → recalcular → CONTROL → guardar
```

## Requisitos de desarrollo

- Python 3.11+
- Windows para probar integración Excel COM
- Microsoft Excel Desktop para la etapa de generación oficial

## Inicio rápido

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
pytest
python -m remasep.main
```
