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
