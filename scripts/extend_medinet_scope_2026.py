"""Regenerate the additional Medinet-derived writable cells for 2026.

Run once against the baseline runtime assets. The historical workbook is never read.
"""

from __future__ import annotations

import csv
import hashlib
import io
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1] / "config" / "runtime_2026"
DETAIL = "'Atenciones - Detalles de citas'"
FP = "stf:dc624775927d4d4d"
DOCTOR = ("CONSULTA MEDICA DE ESPECIALIDAD*",)
NURSE = ("*ENFERMER*",)
SPECIALTIES = {
    55: ("CONSULTA MEDICA DE ESPECIALIDAD EN GENETICA CLINICA*",),
    75: ("CONSULTA MEDICA DE ESPECIALIDAD EN CIRUGIA PLASTICA Y REPARADORA*",),
    89: ("CONSULTA MEDICA DE ESPECIALIDAD EN OTORRINOLARINGOLOGIA*",),
}


def countifs(criteria: tuple[tuple[str, str], ...]) -> str:
    args = ",".join(f'{DETAIL}!{column}:{column},"{value}"' for column, value in criteria)
    return f"COUNTIFS({args})"


def count(types: tuple[str, ...], extra: tuple[tuple[str, str], ...] = ()) -> str:
    return "=" + "+".join(countifs((("O", typ),) + extra) for typ in types)


def age_terms(lo: int, hi: int | None) -> tuple[tuple[str, str], ...]:
    return (("AF", f">={lo}"),) + ((("AF", f"<={hi}"),) if hi is not None else ())


def specialty_count(row: int, types: tuple[str, ...], extra=()) -> str:
    formula = count(types, extra)
    if row == 89:
        # Gantz confirmed inclusion of otorhinolaryngology postoperative
        # controls as consultations. Other specialties and exams stay out.
        formula += "+" + countifs((("M", "Otorrinolaringología"),
                                   ("O", "CONTROL POST OPERATORIO")) + extra)
    return formula


def column_name(number: int) -> str:
    name = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return name


def planned() -> list[tuple[str, str, str]]:
    rows: list[tuple[str, str, str]] = []
    bands_a = ((0, 9), (10, 14), (15, 19), (20, 24), (25, 39),
               (40, 54), (55, 64), (65, None))
    for row, types in ((15, DOCTOR), (16, NURSE)):
        for column, band in zip("DEFGHIJK", bands_a, strict=True):
            rows.append(("REMASEP 01", f"{column}{row}", count(types, age_terms(*band))))
        for column, sex in (("L", "Hombre"), ("M", "Mujer")):
            rows.append(("REMASEP 01", f"{column}{row}", count(types, (("H", sex),))))
    bands_b = tuple((lo, lo + 4) for lo in range(0, 80, 5)) + ((80, None),)
    for row, types in SPECIALTIES.items():
        for column, band in zip("DEFGHIJKLMNOPQRST", bands_b, strict=True):
            rows.append(("REMASEP 01", f"{column}{row}", specialty_count(row, types, age_terms(*band))))
        for column, sex in (("U", "Hombre"), ("V", "Mujer")):
            rows.append(("REMASEP 01", f"{column}{row}", specialty_count(row, types, (("H", sex),))))
    for row, code in ((1251, "1301003"), (1255, "1301021"),
                      (1256, "1301008"), (1258, "1301009"),
                      (1267, "1301045"), (1303, "1303010")):
        # Code prefix in PRESTACION avoids spelling and encoding variations.
        rows.append(("B2 ANEXO", f"C{row}", "=" + countifs((("AA", code + "*"),))))
    coded = "REHABILITACIÓN INTEGRAL DE FONOAUDIÓLOGO - COD.1303011"
    rows.append(("B2 ANEXO", "C1304", count((coded,
        "REHABILITACIÓN FONOAUDIOLOGÍA PRESENCIAL CON GRATUIDAD"), (("K", "Santiago"),))))
    rows.append(("B2 ANEXO", "C1307", count((coded,
        "REHABILITACIÓN FONOAUDIOLOGÍA TELEMEDICINA CON GRATUIDAD"), (("K", "Telemedicina"),))))
    # Gantz feedback (2026-11-26): OD row 58 is instruction in oral hygiene,
    # service code 5010012. Count each eligible appointment once, including
    # when the service appears in a comma-separated list of procedures.
    bands_od = ((0, 1), (2, 3), (4, 4), (5, 5), (6, 6), (7, 7), (8, 8),
                (9, 9), (10, 14), (15, 19), (20, 24), (25, 34), (35, 44),
                (45, 59), (60, 64), (65, 74), (75, None))
    for index, band in enumerate(bands_od):
        for offset, sex in enumerate(("Hombre", "Mujer")):
            criteria = (("AA", "*5010012 -*"), ("H", sex)) + age_terms(*band)
            rows.append(("REMASEP_OD", f"{column_name(5 + index * 2 + offset)}58",
                         "=" + countifs(criteria)))
    assert len(rows) == 119
    return rows


def main() -> None:
    manifest_path = ROOT / "write_manifest.csv"
    catalog_path = ROOT / "metric_catalog.csv"
    with manifest_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        manifest_fields = reader.fieldnames
        manifest = list(reader)
    with catalog_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        catalog_fields = reader.fieldnames
        catalog = list(reader)
    assert manifest_fields and catalog_fields
    # Idempotence: repeat generation updates its own rows without touching baseline.
    manifest = [r for r in manifest if not r["instruction_id"].startswith("medinet:2026:")]
    catalog = [r for r in catalog if not r["metric_id"].startswith("MEDINET::2026::")]
    old_targets = {(r["target_sheet"], r["target_cell"]) for r in manifest}
    for sheet, cell, formula in planned():
        assert (sheet, cell) not in old_targets, (sheet, cell)
        metric_id = f"MEDINET::2026::{sheet.replace(' ', '_')}::{cell}"
        instruction_id = f"medinet:2026:{sheet.replace(' ', '_')}:{cell}"
        otor = sheet == "REMASEP 01" and cell.endswith("89")
        if sheet == "REMASEP_OD":
            signature = f"{sheet} :: EDUCACION SALUD BUCAL :: CODIGO 5010012 :: {cell}"
        elif otor:
            signature = f"{sheet} :: OTORRINO CONSULTA Y CONTROL POST OPERATORIO :: {cell}"
        else:
            signature = f"{sheet} :: MEDINET DIRECT :: {cell}"
        # Virtual namespace prevents new base metrics from shadowing the
        # frozen legacy coordinates used by other metrics' dependencies.
        source_sheet = "MEDINET_DIRECT_REMASEP_OD" if sheet == "REMASEP_OD" else sheet
        catalog.append(dict(metric_id=metric_id, sheet=source_sheet, cell=cell,
                            kind="BASE_AGGREGATION", value_ref_coords="", formula=formula))
        manifest.append(dict(instruction_id=instruction_id, source_metric_id=metric_id,
                             source_semantic_signature=signature, target_sheet=sheet,
                             target_cell=cell, target_semantic_signature=signature,
                             expected_value_type="INTEGER_COUNT", template_fingerprint_id=FP,
                             policy_version="gantz_feedback_20261126" if sheet == "REMASEP_OD" or otor else "medinet_direct_2026.v1",
                             match_status="EXACT_SEMANTIC_MATCH"))
    for path, fields, data, marker in (
        (manifest_path, manifest_fields, manifest, b"medinet:2026:"),
        (catalog_path, catalog_fields, catalog, b"MEDINET::2026::"),
    ):
        # Preserve the frozen CRLF baseline byte-for-byte; new rows use LF to
        # avoid a noisy diff and remain idempotent on repeated generation.
        baseline = b"".join(line for line in path.read_bytes().splitlines(keepends=True)
                            if not line.startswith(marker))
        additions = [row for row in data if (
            row.get("instruction_id", row.get("metric_id", ""))
        ).encode("utf-8").startswith(marker)]
        buffer = io.StringIO(newline="")
        writer = csv.DictWriter(buffer, fieldnames=fields, lineterminator="\n")
        writer.writerows(additions)
        path.write_bytes(baseline + buffer.getvalue().encode("utf-8"))
    bundle_path = ROOT / "bundle.yaml"
    text = bundle_path.read_text(encoding="utf-8")
    import re
    text = re.sub(r"expected_instruction_count: \d+", f"expected_instruction_count: {len(manifest)}", text)
    for path in (manifest_path, catalog_path):
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        text = re.sub(rf"({re.escape(path.name)}: )[0-9a-f]{{64}}", rf"\g<1>{digest}", text)
    bundle_path.write_text(text, encoding="utf-8")
    print(f"instructions={len(manifest)} added={len(planned())}")


if __name__ == "__main__":
    main()
