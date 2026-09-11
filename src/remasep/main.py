"""Punto de entrada de la app (``python -m remasep.main`` / ``REMASEP.exe``).

``--smoke-mode`` (Sprint 3.11, Fase 10) verifica en segundos que el build
puede arrancar — assets internos resueltos, política del writer legible,
``user_data_dir`` escribible — **sin** abrir ninguna ventana ni tocar Excel
COM. Pensado para probar un ``dist/REMASEP/`` recién construido (sobre todo
la variante debug/consola, donde la salida es visible).
"""

from __future__ import annotations

import sys


def main() -> None:
    if "--smoke-mode" in sys.argv[1:]:
        sys.exit(_smoke_check())
    from remasep.ui.main_window import run_app

    run_app()


def _smoke_check() -> int:
    """Sin Qt, sin ventana, sin Excel COM. Devuelve 0 si todo resolvió bien."""
    from remasep import app_paths
    from remasep.logging_setup import configure_logging
    from remasep.services.generation_service import GenerationService
    from remasep.services.runtime_assets import load_runtime_bundle

    log = configure_logging()

    def report(line: str) -> None:
        print(line)
        log.info(line)

    report("=== REMASEP Automation — smoke check ===")
    report(f"frozen: {app_paths.is_frozen()}")
    report(f"resource_root: {app_paths.resource_root()}")

    bundle = load_runtime_bundle()
    report(
        f"runtime_2026 OK: version={bundle.version} "
        f"instrucciones={len(bundle.write_instructions)} "
        f"fingerprint={bundle.template_fingerprint_id}"
    )

    service = GenerationService(
        config_dir=app_paths.excel_writer_config_dir(),
        artifacts_dir=app_paths.diagnostics_dir(),
    )
    policy = service._policy()
    report(f"excel_writer_2026 OK: version={policy.get('version')}")

    report(f"user_data_dir: {app_paths.user_data_dir()}")
    report(f"logs_dir: {app_paths.logs_dir()}")
    report(f"diagnostics_dir: {app_paths.diagnostics_dir()}")
    report("SMOKE OK")
    return 0


if __name__ == "__main__":
    main()
