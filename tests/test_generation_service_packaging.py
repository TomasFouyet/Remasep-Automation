"""Sprint 3.11 — la UI inyecta paths de app empaquetada al construir el
``GenerationService`` por defecto (Fase 2: config interno + diagnósticos en
``user_data_dir``, nunca en el *cwd*). El CLI (``GenerationService()`` sin
argumentos) no se toca: sigue usando sus defaults relativos de siempre.
"""

from __future__ import annotations

from pathlib import Path

from remasep import app_paths
from remasep.services.common import Period
from remasep.ui import workers


def test_run_generation_injects_packaged_config_and_diagnostics_dirs(
    tmp_path, monkeypatch, make_medinet
):
    from datetime import date

    monkeypatch.setenv("REMASEP_APP_DATA_DIR", str(tmp_path / "appdata"))
    monkeypatch.chdir(tmp_path)  # cwd irrelevante: nada debe depender de él

    captured = {}
    real_service_cls = workers.GenerationService

    class SpyService(real_service_cls):
        def __init__(self, *, config_dir, artifacts_dir):
            captured["config_dir"] = Path(config_dir)
            captured["artifacts_dir"] = Path(artifacts_dir)
            super().__init__(config_dir=config_dir, artifacts_dir=artifacts_dir)

    monkeypatch.setattr(workers, "GenerationService", SpyService)

    medinet = make_medinet([
        [date(2026, 7, 10), date(2016, 1, 1), "Mujer", "Centro", "Ortodoncia",
         "CONSULTA", "prestacion", "Atendido", "Presencial", ""]
    ])
    workers.run_generation(medinet, Period(7, 2026), "plantilla.xlsm", None)

    assert captured["config_dir"] == app_paths.excel_writer_config_dir()
    assert captured["artifacts_dir"] == app_paths.diagnostics_dir()
    # nunca junto al cwd de la ejecución
    assert not str(captured["artifacts_dir"]).startswith(str(tmp_path / "artifacts"))
    assert str(captured["artifacts_dir"]).startswith(str(tmp_path / "appdata"))
    assert not (tmp_path / "artifacts").exists()
