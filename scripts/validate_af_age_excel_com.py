"""Validación independiente de `legacy_age_years` (AF) contra Microsoft Excel real.

Diagnóstico de sólo lectura, ejecutado una sola vez para cerrar la duda sobre si
`legacy_age_years` (``src/remasep/services/legacy_transform.py``) reproduce
exactamente la fórmula legacy del workbook::

    =IF(FECHA_NACIMIENTO > DIA_CITA, 0, DATEDIF(FECHA_NACIMIENTO, DIA_CITA, "Y"))

La referencia es Microsoft Excel real, no los valores AF cacheados del workbook
``GENERACION DATOS REMASEP.xlsx`` (esos ya se sabe que están obsoletos —
ver `legacy_derived_equivalence`).

**No reimplementa nada**: reutiliza `processing_scope_frame` (Sprint 2.1) para
obtener los registros estructuralmente válidos de cada período y
`legacy_age_years` (motor productivo) para el valor esperado en Python.

Puente Excel COM
-----------------
Este script NO usa ``pywin32`` (a diferencia de ``remasep.adapters.excel_com``,
que sí lo requiere para el flujo productivo en Windows nativo). Es un
diagnóstico de una sola vez, así que usa un puente más portable: invoca
``powershell.exe`` como subproceso (funciona tanto en Windows nativo como desde
WSL vía interop, que es el entorno real de desarrollo de este proyecto) y le
pasa los pares de fechas por **stdin** — nunca se escribe un CSV ni ningún
archivo intermedio a disco, así que no hay fechas de nacimiento individuales
que limpiar ni que puedan filtrarse a un log.

El script de PowerShell:

- crea una instancia Excel **propia** (``New-Object -ComObject``), nunca se
  conecta a una instancia existente (``GetActiveObject``);
- añade un workbook en memoria (``Workbooks.Add()``), nunca abre ni guarda un
  archivo;
- escribe únicamente las fechas recibidas por stdin y la fórmula de
  comparación;
- calcula, lee los resultados y los imprime por stdout (sólo enteros: la edad
  calculada, nunca las fechas);
- cierra el workbook sin guardar y hace ``Quit()`` de su propia instancia en un
  bloque ``finally`` (nunca mata procesos).

Uso::

    python scripts/validate_af_age_excel_com.py
    python scripts/validate_af_age_excel_com.py --months 7 --year 2026
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd

from remasep.core.errors import RemasepError
from remasep.services.common import Period
from remasep.services.legacy_transform import legacy_age_years
from remasep.services.medinet_analysis import processing_scope_frame

DEFAULT_EXPORT = "data/local/detalle_citas - 2026-09-07T123630.940.xlsx"
YEAR_IN_SCOPE = 2026
MONTHS_IN_SCOPE = (4, 5, 6, 7)

DatePair = tuple[date, date]

# ---------------------------------------------------------------------------
# Extracción de pares FECHA_NACIMIENTO / DIA_CITA (sin reimplementar el scope)
# ---------------------------------------------------------------------------


def extract_date_pairs(
    export_path: str | Path,
    months: Iterable[int] = MONTHS_IN_SCOPE,
    year: int = YEAR_IN_SCOPE,
) -> list[DatePair]:
    """Pares únicos ``(FECHA_NACIMIENTO, DIA_CITA)`` de los registros
    estructuralmente válidos de ``months``/``year``.

    Reutiliza ``processing_scope_frame`` (motor productivo, sin ESTADO) para
    cada mes — el mismo subconjunto de registros que alimenta AF en producción.
    Los registros sin ``FECHA_NACIMIENTO`` (permitido: el campo es opcional en
    el scope estructural) se excluyen: `legacy_age_years` no puede evaluar la
    fórmula sin ambas fechas y no hay nada que comparar contra Excel.
    """
    pairs: set[DatePair] = set()
    for month in months:
        frame = processing_scope_frame(export_path, Period(month, year))
        subset = frame[["FECHA_NACIMIENTO", "DIA_CITA"]].dropna()
        for fnac, dia in zip(subset["FECHA_NACIMIENTO"], subset["DIA_CITA"], strict=True):
            pairs.add((pd.Timestamp(fnac).date(), pd.Timestamp(dia).date()))
    return sorted(pairs)


def python_expected_ages(pairs: list[DatePair]) -> dict[DatePair, int | None]:
    """Valor esperado por par según `legacy_age_years` (motor productivo)."""
    return {pair: legacy_age_years(pair[0], pair[1]) for pair in pairs}


# ---------------------------------------------------------------------------
# Puente Excel COM (subproceso PowerShell, propio y aislado)
# ---------------------------------------------------------------------------

# Excel COM en su propia instancia; nunca se conecta a una existente
# (GetActiveObject), nunca abre/guarda un archivo. Recibe pares por stdin
# ("yyyy-MM-dd,yyyy-MM-dd" por línea) y sólo imprime enteros (la edad
# calculada) por stdout, uno por línea, en el mismo orden.
_POWERSHELL_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$lines = @()
while (($line = [Console]::In.ReadLine()) -ne $null) {
    if ($line.Trim().Length -gt 0) { $lines += $line }
}
$n = $lines.Count
if ($n -eq 0) { exit 0 }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
$ws = $null
try {
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    $birthArr = New-Object 'object[,]' $n,1
    $serviceArr = New-Object 'object[,]' $n,1
    for ($i = 0; $i -lt $n; $i++) {
        $parts = $lines[$i] -split ','
        $birthArr[$i,0] = [DateTime]::ParseExact($parts[0], 'yyyy-MM-dd', $null)
        $serviceArr[$i,0] = [DateTime]::ParseExact($parts[1], 'yyyy-MM-dd', $null)
    }
    $ws.Range("A1", "A$n").Value2 = $birthArr
    $ws.Range("B1", "B$n").Value2 = $serviceArr
    $ws.Range("C1", "C$n").FormulaR1C1 = '=IF(RC[-2]>RC[-1],0,DATEDIF(RC[-2],RC[-1],"Y"))'
    $excel.Calculate()

    $resultRange = $ws.Range("C1", "C$n")
    if ($n -eq 1) {
        Write-Output ([string]$resultRange.Value2)
    } else {
        $values = $resultRange.Value2
        for ($i = 1; $i -le $n; $i++) {
            Write-Output ([string]$values[$i,1])
        }
    }
}
finally {
    if ($wb -ne $null) { $wb.Close($false) }
    $excel.Quit()
    if ($ws -ne $null) { [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null }
    if ($wb -ne $null) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""


def find_powershell() -> str | None:
    return shutil.which("powershell.exe") or shutil.which("pwsh")


def excel_com_available(powershell_exe: str | None = None) -> bool:
    """Prueba mínima: ¿se puede crear y cerrar una instancia Excel propia?

    No escribe nada; crea la instancia y la cierra de inmediato. Si falla por
    cualquier motivo (Excel no instalado, COM no disponible, etc.) devuelve
    ``False`` sin lanzar.
    """
    exe = powershell_exe or find_powershell()
    if exe is None:
        return False
    probe = (
        "$e = New-Object -ComObject Excel.Application; "
        "$e.Quit(); "
        "[Runtime.InteropServices.Marshal]::ReleaseComObject($e) | Out-Null"
    )
    try:
        result = subprocess.run(
            [exe, "-NoProfile", "-NonInteractive", "-Command", probe],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return result.returncode == 0


def _powershell_with_script(exe: str) -> list[str]:
    """Argumentos para ejecutar `_POWERSHELL_SCRIPT` con los datos por stdin.

    El script se pasa codificado en base64 (``-EncodedCommand``): PowerShell no
    permite mezclar de forma fiable un script inline largo con datos por stdin
    en un único ``-Command "-"``.
    """
    import base64

    encoded = base64.b64encode(_POWERSHELL_SCRIPT.encode("utf-16-le")).decode("ascii")
    return [exe, "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded]


def run_excel_com_ages(
    pairs: list[DatePair], powershell_exe: str | None = None
) -> list[int | None]:
    """Calcula la fórmula legacy en Excel real para ``pairs``, en el mismo orden.

    ``None`` en la posición ``i`` significa que Excel no pudo evaluar ese par
    (línea de salida ausente o no numérica) — se reporta como error, nunca se
    fabrica un valor.
    """
    exe = powershell_exe or find_powershell()
    if exe is None:
        raise RemasepError(
            "powershell.exe no disponible: la validación Excel COM requiere "
            "Windows (nativo o WSL con interop) con Microsoft Excel instalado."
        )
    if not pairs:
        return []

    stdin_payload = "\n".join(f"{b.isoformat()},{s.isoformat()}" for b, s in pairs) + "\n"
    try:
        result = subprocess.run(
            _powershell_with_script(exe),
            input=stdin_payload,
            capture_output=True,
            text=True,
            timeout=600,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise RemasepError("la sesión Excel COM excedió el tiempo límite") from exc

    if result.returncode != 0:
        raise RemasepError(
            f"la sesión Excel COM terminó con error (code={result.returncode}): "
            f"{result.stderr.strip()}"
        )

    out_lines = [line for line in result.stdout.splitlines() if line.strip() != ""]
    ages: list[int | None] = []
    for i in range(len(pairs)):
        if i >= len(out_lines):
            ages.append(None)
            continue
        try:
            ages.append(int(float(out_lines[i].strip())))
        except ValueError:
            ages.append(None)
    return ages


# ---------------------------------------------------------------------------
# Comparación (sin PII: nunca se conservan fechas en el resultado)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class MismatchDetail:
    index: int
    python_value: int | None
    excel_value: int | None


@dataclass(frozen=True)
class AFValidationResult:
    total: int
    match: int
    mismatch: int
    errors: int
    mismatches: list[MismatchDetail] = field(default_factory=list)

    @property
    def equivalence_rate(self) -> float | None:
        return (self.match / self.total) if self.total else None


def compare_ages(
    pairs: list[DatePair],
    python_ages: dict[DatePair, int | None],
    excel_ages: list[int | None],
) -> AFValidationResult:
    """Compara ``python_ages`` (por par) contra ``excel_ages`` (alineado por
    índice a ``pairs``). Nunca incluye fechas en el resultado — sólo índices y
    valores enteros."""
    if len(pairs) != len(excel_ages):
        raise RemasepError(
            f"pairs ({len(pairs)}) y excel_ages ({len(excel_ages)}) deben tener el mismo largo"
        )
    match = 0
    errors = 0
    mismatches: list[MismatchDetail] = []
    for i, pair in enumerate(pairs):
        py_value = python_ages[pair]
        excel_value = excel_ages[i]
        if excel_value is None:
            errors += 1
            continue
        if py_value == excel_value:
            match += 1
        else:
            mismatches.append(MismatchDetail(i, py_value, excel_value))
    return AFValidationResult(
        total=len(pairs),
        match=match,
        mismatch=len(mismatches),
        errors=errors,
        mismatches=mismatches,
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="validate_af_age_excel_com.py",
        description=(
            "Valida legacy_age_years (AF) contra Microsoft Excel real vía COM. "
            "Diagnóstico de sólo lectura; no modifica legacy_age_years."
        ),
    )
    parser.add_argument("--export", default=DEFAULT_EXPORT)
    parser.add_argument("--months", type=int, nargs="+", default=list(MONTHS_IN_SCOPE))
    parser.add_argument("--year", type=int, default=YEAR_IN_SCOPE)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    pairs = extract_date_pairs(args.export, args.months, args.year)
    print(f"Pares únicos (FECHA_NACIMIENTO, DIA_CITA) a validar: {len(pairs)}")

    if not excel_com_available():
        print(
            "ERROR: Excel COM no disponible en este entorno "
            "(se requiere powershell.exe + Microsoft Excel instalado).",
            file=sys.stderr,
        )
        return 2

    python_ages = python_expected_ages(pairs)
    excel_ages = run_excel_com_ages(pairs)
    outcome = compare_ages(pairs, python_ages, excel_ages)

    print(f"total={outcome.total} MATCH={outcome.match} MISMATCH={outcome.mismatch} "
          f"ERRORS={outcome.errors}")
    rate = outcome.equivalence_rate
    print(f"equivalence_rate={rate:.4%}" if rate is not None else "equivalence_rate=N/A")

    if outcome.mismatches:
        print("Mismatches (índice, python, excel) — sin fechas:")
        for m in outcome.mismatches:
            print(f"  #{m.index}: python={m.python_value} excel={m.excel_value}")

    return 0 if (outcome.mismatch == 0 and outcome.errors == 0) else 1


if __name__ == "__main__":
    raise SystemExit(main())
