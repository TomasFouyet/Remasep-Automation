"""Smoke test estructural sobre un ``dist/REMASEP/`` ya construido (Sprint 3.11,
Fase 10). Confirma que el build de PyInstaller (ONEDIR) trae lo necesario y
**no** trae nada de desarrollo — sin necesitar Windows para las
comprobaciones de archivos/rutas.

Uso::

    python scripts/build_smoke_test.py
    python scripts/build_smoke_test.py --dist dist/REMASEP

Sólo inspecciona archivos/rutas — nunca lanza procesos (este repo no importa
``subprocess`` en ``src/``/``scripts/``: ver
``tests/test_excel_writer.py::test_no_process_killing_in_codebase``). Para
verificar que el ``.exe`` arranca de verdad, en Windows y a mano:

    dist\\REMASEP\\REMASEP-debug.exe --smoke-mode

(``remasep.main`` resuelve sus assets internos y sale sin abrir ninguna
ventana ni tocar Excel COM — ver docs/WINDOWS_PACKAGING.md, Fase 10.)
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from pathlib import Path

# Deben existir dentro del dist (Fase 3/10).
REQUIRED_RELATIVE_PATHS = (
    "REMASEP.exe",
    "_internal/config/runtime_2026/bundle.yaml",
    "_internal/config/runtime_2026/write_manifest.csv",
    "_internal/config/runtime_2026/metric_catalog.csv",
    "_internal/config/excel_writer_2026/policy.yaml",
    "_internal/config/excel_writer_2026/control_map.yaml",
)

# NO deben aparecer en ningún nombre de archivo/carpeta del dist (Fase 3).
FORBIDDEN_NEEDLES = (
    "data/local",
    "generacion datos remasep",
    "detalle_citas",  # nombre típico de un export real de Medinet
)

# NO deben existir como carpeta en ningún punto del árbol.
FORBIDDEN_DIR_NAMES = ("artifacts", "tests", ".git", ".venv", ".venv-win", "outputs")


@dataclass
class SmokeResult:
    ok: bool
    checks: list[tuple[str, bool, str]] = field(default_factory=list)


def check_dist(dist: Path) -> SmokeResult:
    """Sólo lectura: nunca borra ni modifica nada del dist."""
    result = SmokeResult(ok=True)

    def add(name: str, ok: bool, detail: str = "") -> None:
        result.checks.append((name, ok, detail))
        if not ok:
            result.ok = False

    add("el directorio dist existe", dist.is_dir(), str(dist))
    if not dist.is_dir():
        return result

    for rel in REQUIRED_RELATIVE_PATHS:
        p = dist / rel
        add(f"presente: {rel}", p.is_file(), str(p))

    all_paths = [str(p).replace("\\", "/") for p in dist.rglob("*")]
    haystack = "\n".join(all_paths).lower()

    for needle in FORBIDDEN_NEEDLES:
        add(f'ninguna ruta contiene "{needle}"', needle not in haystack)

    present_forbidden = sorted(
        {name for name in FORBIDDEN_DIR_NAMES if any(p.name == name for p in dist.rglob(name))}
    )
    add(
        "sin carpetas de desarrollo (artifacts/tests/.git/.venv*/outputs)",
        not present_forbidden,
        ", ".join(present_forbidden),
    )

    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dist", default="dist/REMASEP", help="carpeta dist a inspeccionar")
    args = parser.parse_args(argv)

    dist = Path(args.dist)
    result = check_dist(dist)
    for name, ok, detail in result.checks:
        mark = "OK   " if ok else "FALLA"
        suffix = f" — {detail}" if detail and not ok else ""
        print(f"[{mark}] {name}{suffix}")

    print()
    print("RESULTADO:", "OK" if result.ok else "FALLA")
    if result.ok:
        print(
            "Para confirmar que arranca de verdad (Windows, a mano):\n"
            f"  {dist}\\REMASEP-debug.exe --smoke-mode"
        )
    return 0 if result.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
