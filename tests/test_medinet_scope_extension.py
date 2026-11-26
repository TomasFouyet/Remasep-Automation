"""Cross-sheet reconciliation for the Medinet-only 2026 extension."""

from datetime import date

from remasep.services.common import Period
from remasep.services.production_pipeline import build_production_pending_writes


def test_new_sections_respect_age_sex_branch_and_estado(make_medinet):
    def row(birth, sex, specialty, kind, branch="Santiago", status="Atendido", service=""):
        return [date(2026, 9, 15), birth, sex, branch, specialty, kind,
                service, status, "Presencial", ""]

    coded = "REHABILITACIÓN INTEGRAL DE FONOAUDIÓLOGO - COD.1303011"
    rows = [
        row(date(2018, 9, 15), "Mujer", "Cirugía",
            "CONSULTA MEDICA DE ESPECIALIDAD EN CIRUGIA PLASTICA Y REPARADORA - COD.0101316"),
        row(date(2014, 9, 14), "Hombre", "Enfermería", "ATENCIÓN DE ENFERMERÍA"),
        row(date(2018, 9, 15), "Mujer", "Cirugía",
            "CONSULTA MEDICA DE ESPECIALIDAD EN CIRUGIA PLASTICA Y REPARADORA - COD.0101316",
            status="Agendado"),
        row(date(2000, 1, 1), "Mujer", "Fonoaudiología", coded),
        row(date(2000, 1, 1), "Mujer", "Fonoaudiología", coded, "Telemedicina"),
        row(date(2000, 1, 1), "Mujer", "Fonoaudiología",
            "REHABILITACIÓN FONOAUDIOLOGÍA PRESENCIAL CON GRATUIDAD"),
        row(date(2000, 1, 1), "Mujer", "Fonoaudiología",
            "REHABILITACIÓN FONOAUDIOLOGÍA TELEMEDICINA CON GRATUIDAD", "Telemedicina"),
        row(date(2000, 1, 1), "Mujer", "Otorrinolaringología",
            "NASOFARINGOLARINGOFIBROSCOPIA - COD.1301003", service="1301003 - NASOFARINGOLARINGOFIBROSCOPIA"),
    ]
    result = build_production_pending_writes(make_medinet(rows), Period(9, 2026))
    values = {(w.target_sheet, w.target_cell): w.value for w in result.pending_writes}
    assert result.completeness.ok and len(values) == 1241
    assert result.scope.processing_scope_records == 7
    assert [values["REMASEP 01", cell] for cell in ("D15", "L15", "M15", "E16", "L16")] == [1, 0, 1, 1, 1]
    assert values["B2 ANEXO", "C1304"] == 2
    assert values["B2 ANEXO", "C1307"] == 2
    assert values["B2 ANEXO", "C1251"] == 1


def test_oral_hygiene_education_uses_confirmed_code_once_per_eligible_cita(make_medinet):
    def row(birth, sex, service, estado="Atendido"):
        return [date(2026, 10, 15), birth, sex, "Santiago", "Ortodoncia",
                "SESIÓN TTO ORTODONCIA", service, estado, "Fonasa A", ""]

    code = "5010012 - INSTRUCCIÓN DE HIGIENE ORAL ODONTOLOGÍA"
    rows = [
        row(date(2000, 1, 1), "Mujer", "5001016 - FLUOR, " + code),
        row(date(2014, 1, 1), "Hombre", code + ", " + code),
        row(date(2000, 1, 1), "Mujer", code, "Cancelado"),
        row(date(2000, 1, 1), "Mujer", "5010011 - PROFILAXIS"),
    ]
    result = build_production_pending_writes(make_medinet(rows), Period(10, 2026))
    values = {(w.target_sheet, w.target_cell): w.value for w in result.pending_writes}
    assert len(values) == 1241 and result.completeness.ok
    assert values["REMASEP_OD", "AB58"] == 1  # Mujer, 25–34 años
    assert values["REMASEP_OD", "U58"] == 1   # Hombre, 10–14 años
    additional = [w for w in result.pending_writes
                  if w.target_sheet == "REMASEP_OD" and w.source_metric_id.startswith("MEDINET::2026::")]
    assert len(additional) == 34
    assert sum(w.value for w in additional) == 2


def test_otor_consultations_include_only_otor_postoperative_controls(make_medinet):
    def row(kind, specialty="Otorrinolaringología", *, state="Atendido", month=10):
        return [date(2026, month, 15), date(2014, 1, 1), "Hombre", "Santiago",
                specialty, kind, "", state, "Fonasa A", ""]

    rows = [
        row("CONSULTA MEDICA DE ESPECIALIDAD EN OTORRINOLARINGOLOGIA - COD.0101205"),
        row("CONTROL POST OPERATORIO"),
        row("CONTROL POST OPERATORIO", specialty="Cirugía"),
        row("IMPEDANCIOMETRÍA - COD.1301009"),
        row("CONTROL POST OPERATORIO", state="Cancelado"),
        row("CONTROL POST OPERATORIO", month=9),
    ]
    result = build_production_pending_writes(make_medinet(rows), Period(10, 2026))
    values = {(w.target_sheet, w.target_cell): w.value for w in result.pending_writes}
    assert result.completeness.ok and len(values) == 1241
    assert values["REMASEP 01", "F89"] == 2  # Hombre, 10–14 años
    assert values["REMASEP 01", "U89"] == 2
    assert values["REMASEP 01", "V89"] == 0
    assert sum(values["REMASEP 01", f"{col}89"] for col in "DEFGHIJKLMNOPQRST") == 2
    # This confirmation concerns Section B; it does not extend Control de Salud.
    assert sum(values["REMASEP 01", f"{col}15"] for col in "DEFGHIJK") == 1


def test_nasofibroscopy_includes_contingencies_once_and_keeps_scope(make_medinet):
    coded = "NASOFARINGOLARINGOFIBROSCOPIA - COD.1301003"
    contingency = "ATENCION DE CONTINGENCIA NASOFARINGOLARINGOFIBROSCOPIA"
    service = "1301003 - NASOFARINGOLARINGOFIBROSCOPIA"

    def row(kind, procedure="", *, state="Atendido", month=10):
        return [date(2026, month, 15), date(2000, 1, 1), "Mujer", "Santiago",
                "Otorrinolaringología", kind, procedure, state, "Fonasa A", ""]

    rows = [
        row(coded, service),
        row(contingency),
        row(contingency, service),  # Both criteria: one appointment, one count.
        row(contingency, "1301009 - IMPEDANCIOMETRÍA"),
        row("ATENCION DE CONTINGENCIA EXAMENES AUDITIVOS"),
        row(contingency, state="Cancelado"),
        row(contingency, month=9),
    ]
    result = build_production_pending_writes(make_medinet(rows), Period(10, 2026))
    values = {(w.target_sheet, w.target_cell): w.value for w in result.pending_writes}
    assert result.completeness.ok and len(values) == 1241
    assert values["B2 ANEXO", "C1251"] == 4
    assert values["B2 ANEXO", "C1258"] == 1  # Existing procedure criterion retained.


def test_psychology_support_meetings_stay_in_their_branch(make_medinet, legacy_psychology_bundle):
    consultation = "CONSULTA PSICOLOGO CLINICO (SESIONES 45') - COD.0902001"

    def row(kind, branch, *, specialty="Psicología", state="Atendido", month=10):
        return [date(2026, month, 15), date(2000, 1, 1), "Mujer", branch,
                specialty, kind, "", state, "Fonasa A", ""]

    rows = [
        row(consultation, "Santiago"), row(consultation, "Telemedicina"),
        row("REUNIÓN APOYO", "Santiago"), row("REUNIÓN APOYO", "Telemedicina"),
        row("REUNIÓN APOYO", "Santiago", specialty="Fonoaudiología"),
        row("REUNIÓN APOYO", "Otra sucursal"),
        row("REUNIÓN APOYO", "Santiago", state="Cancelado"),
        row("REUNIÓN APOYO", "Telemedicina", month=9),
    ]
    path = make_medinet(rows)
    current = build_production_pending_writes(path, Period(10, 2026))
    assert current.completeness.ok and len(current.pending_writes) == 1241
    values = {(w.target_sheet, w.target_cell): w.value for w in current.pending_writes}
    assert values["B2 ANEXO", "C953"] == values["B2 ANEXO", "C957"] == 2
    archived = build_production_pending_writes(path, Period(10, 2026), bundle=legacy_psychology_bundle)
    old_values = {(w.target_sheet, w.target_cell): w.value for w in archived.pending_writes}
    assert old_values["B2 ANEXO", "C953"] == old_values["B2 ANEXO", "C957"] == 1
    changed = {target for target in values if values[target] != old_values[target]}
    assert changed == {("B2 ANEXO", "C953"), ("B2 ANEXO", "C957")}


def test_asset_generator_is_idempotent_and_preserves_archived_catalog(tmp_path, monkeypatch):
    import importlib.util
    import shutil
    from pathlib import Path

    from remasep.services.runtime_assets import load_runtime_bundle

    root = Path(__file__).resolve().parents[1]
    spec = importlib.util.spec_from_file_location("extend_scope", root / "scripts/extend_medinet_scope_2026.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    target = tmp_path / "runtime_2026"
    shutil.copytree(root / "config/runtime_2026", target)
    monkeypatch.setattr(module, "ROOT", target)
    module.main()
    first = {p.name: p.read_bytes() for p in target.iterdir() if p.is_file()}
    module.main()
    assert first == {p.name: p.read_bytes() for p in target.iterdir() if p.is_file()}
    bundle = load_runtime_bundle(target)
    assert len(bundle.write_instructions) == 1241
    for cell, legacy_id in module.OVERRIDES.items():
        active = next(w for w in bundle.write_instructions
                      if (w.target_sheet, w.target_cell) == cell)
        assert active.source_metric_id.startswith("MEDINET::2026::")
        assert legacy_id in bundle.formula_index
