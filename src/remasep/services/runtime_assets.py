"""Assets de runtime versionados con la aplicación (Sprint 3.8).

El *path* de producción **no** abre ``GENERACION DATOS REMASEP.xlsx`` ni lee
nada bajo ``artifacts/``. Todo el conocimiento estructural ya validado
(manifiesto de escritura, catálogo de fórmulas por métrica, contrato de la hoja
de detalle, política del cero, fingerprint de plantilla) vive congelado y
auditado bajo ``config/runtime_2026/``.

Estos assets contienen **sólo** reglas, mappings, IDs y metadatos estructurales:
ninguna fila de Medinet, ningún nombre/RUT/fecha de nacimiento, ningún valor
mensual real.

Un asset ausente o corrupto produce :class:`RuntimeAssetError` (motivo
``RUNTIME_ASSET_MISSING`` / ``RUNTIME_ASSET_INCOMPATIBLE`` /
``RUNTIME_ASSET_INVALID``), nunca un ``FileNotFoundError`` crudo.
"""

from __future__ import annotations

import csv
import hashlib
import sys
from collections.abc import Callable, Mapping
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from remasep.core.errors import RemasepError
from remasep.services.metric_value_producer import LegacyFormulaSpec

RUNTIME_DIR_NAME = "runtime_2026"
_BUNDLE_FILE = "bundle.yaml"

RUNTIME_ASSET_MISSING = "RUNTIME_ASSET_MISSING"
RUNTIME_ASSET_INCOMPATIBLE = "RUNTIME_ASSET_INCOMPATIBLE"
RUNTIME_ASSET_INVALID = "RUNTIME_ASSET_INVALID"

_USER_MESSAGES = {
    RUNTIME_ASSET_MISSING: (
        "La instalación no contiene los archivos internos necesarios para "
        "generar el informe."
    ),
    RUNTIME_ASSET_INCOMPATIBLE: (
        "Los archivos internos de la instalación no son compatibles con esta "
        "plantilla REMASEP."
    ),
    RUNTIME_ASSET_INVALID: (
        "Los archivos internos de la instalación están dañados o incompletos."
    ),
}

_MANIFEST_COLUMNS = (
    "instruction_id",
    "source_metric_id",
    "source_semantic_signature",
    "target_sheet",
    "target_cell",
    "target_semantic_signature",
    "expected_value_type",
    "template_fingerprint_id",
    "policy_version",
    "match_status",
)
_CATALOG_COLUMNS = ("metric_id", "sheet", "cell", "kind", "value_ref_coords", "formula")

_ACCEPTED_ZERO_POLICIES = ("WRITE_ZERO", "FORM_SPECIFIC")


class RuntimeAssetError(RemasepError):
    """Falta / corrupción / incompatibilidad de un asset de runtime."""

    def __init__(self, code: str, detail: str = "") -> None:
        self.code = code
        self.detail = detail
        message = _USER_MESSAGES.get(code, "Error de asset de runtime.")
        super().__init__(f"{code}: {message}" + (f" ({detail})" if detail else ""))

    @property
    def user_message(self) -> str:
        return _USER_MESSAGES.get(self.code, "Error de asset de runtime.")


# ---------------------------------------------------------------------------
# Resolución de paths (repo y, a futuro, app empaquetada con PyInstaller)
# ---------------------------------------------------------------------------


def resolve_runtime_root(explicit: str | Path | None = None) -> Path:
    """Localiza ``config/runtime_2026/`` sin depender del *current working dir*.

    Orden: ``explicit`` → bundle PyInstaller (``sys._MEIPASS``) → subiendo desde
    este módulo hasta encontrar ``config/<RUNTIME_DIR_NAME>``.
    """
    if explicit is not None:
        root = Path(explicit)
        if root.is_dir():
            return root
        raise RuntimeAssetError(RUNTIME_ASSET_MISSING, f"no existe el directorio {root}")

    candidates: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "config" / RUNTIME_DIR_NAME)
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidates.append(parent / "config" / RUNTIME_DIR_NAME)

    for cand in candidates:
        if (cand / _BUNDLE_FILE).is_file():
            return cand
    raise RuntimeAssetError(
        RUNTIME_ASSET_MISSING,
        f"no se encontró config/{RUNTIME_DIR_NAME}/{_BUNDLE_FILE}",
    )


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class WriteInstructionRow:
    """Contrato semántico de una instrucción de escritura (sin PII, sin valores)."""

    instruction_id: str
    source_metric_id: str
    source_semantic_signature: str
    target_sheet: str
    target_cell: str
    target_semantic_signature: str
    expected_value_type: str
    template_fingerprint_id: str
    policy_version: str
    match_status: str


@dataclass
class RuntimeBundle:
    version: str
    root: str
    template_fingerprint_id: str
    detail_sheet: str
    detail_columns_map: dict[str, str]
    criterion_values: dict[tuple[str, str], object]
    write_instructions: list[WriteInstructionRow]
    formula_index: dict[str, LegacyFormulaSpec]
    expected_instruction_count: int
    zero_write_policy: str
    estado_filter_status: str
    asset_sha256: dict[str, str] = field(default_factory=dict)

    # -- interfaces que consume el productor -------------------------
    @property
    def instruction_ids(self) -> list[str]:
        return [w.instruction_id for w in self.write_instructions]

    def expected_value_types(self) -> dict[str, str]:
        return {w.source_metric_id: w.expected_value_type for w in self.write_instructions}

    def resolver_for(self, sheet: str) -> Callable[[str], object]:
        values = self.criterion_values

        def resolve(ref: str) -> object:
            return values.get((sheet, ref.replace("$", "").upper()))

        return resolve

    def resolver_by_sheet(self) -> dict[str, Callable[[str], object]]:
        sheets = {spec.sheet for spec in self.formula_index.values()}
        return {sheet: self.resolver_for(sheet) for sheet in sheets}


# ---------------------------------------------------------------------------
# Carga + validación
# ---------------------------------------------------------------------------


def _load_yaml(path: Path) -> dict:
    try:
        return yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError) as exc:
        raise RuntimeAssetError(RUNTIME_ASSET_INVALID, f"{path.name}: {exc}") from exc


def _require(path: Path) -> Path:
    if not path.is_file():
        raise RuntimeAssetError(RUNTIME_ASSET_MISSING, f"falta {path.name}")
    return path


def _read_manifest(path: Path) -> list[WriteInstructionRow]:
    _require(path)
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        header = tuple(reader.fieldnames or ())
        if header != _MANIFEST_COLUMNS:
            raise RuntimeAssetError(
                RUNTIME_ASSET_INVALID,
                f"{path.name}: esquema inesperado {header}",
            )
        rows = [
            WriteInstructionRow(
                instruction_id=r["instruction_id"],
                source_metric_id=r["source_metric_id"],
                source_semantic_signature=r["source_semantic_signature"],
                target_sheet=r["target_sheet"],
                target_cell=r["target_cell"],
                target_semantic_signature=r["target_semantic_signature"],
                expected_value_type=r["expected_value_type"],
                template_fingerprint_id=r["template_fingerprint_id"],
                policy_version=r["policy_version"],
                match_status=r["match_status"],
            )
            for r in reader
        ]
    if not rows:
        raise RuntimeAssetError(RUNTIME_ASSET_INVALID, f"{path.name}: sin filas")
    return rows


def _read_catalog(path: Path) -> dict[str, LegacyFormulaSpec]:
    _require(path)
    index: dict[str, LegacyFormulaSpec] = {}
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        header = tuple(reader.fieldnames or ())
        if header != _CATALOG_COLUMNS:
            raise RuntimeAssetError(
                RUNTIME_ASSET_INVALID, f"{path.name}: esquema inesperado {header}"
            )
        for r in reader:
            mid = r["metric_id"]
            if mid in index:
                raise RuntimeAssetError(
                    RUNTIME_ASSET_INVALID, f"{path.name}: metric_id duplicado {mid}"
                )
            coords = tuple(c for c in (r["value_ref_coords"] or "").split("|") if c)
            index[mid] = LegacyFormulaSpec(
                source_metric_id=mid,
                sheet=r["sheet"],
                cell=r["cell"],
                kind=r["kind"],
                formula=r["formula"],
                value_ref_coords=coords,
            )
    if not index:
        raise RuntimeAssetError(RUNTIME_ASSET_INVALID, f"{path.name}: catálogo vacío")
    return index


def _validate(bundle: RuntimeBundle) -> None:
    ids = bundle.instruction_ids
    if len(ids) != len(set(ids)):
        raise RuntimeAssetError(RUNTIME_ASSET_INVALID, "instruction_id duplicados")
    cells = [(w.target_sheet, w.target_cell) for w in bundle.write_instructions]
    if len(cells) != len(set(cells)):
        raise RuntimeAssetError(RUNTIME_ASSET_INVALID, "celdas destino duplicadas")
    if len(bundle.write_instructions) != bundle.expected_instruction_count:
        raise RuntimeAssetError(
            RUNTIME_ASSET_INVALID,
            f"se esperaban {bundle.expected_instruction_count} instrucciones, "
            f"hay {len(bundle.write_instructions)}",
        )

    fps = {w.template_fingerprint_id for w in bundle.write_instructions}
    if fps != {bundle.template_fingerprint_id}:
        raise RuntimeAssetError(
            RUNTIME_ASSET_INCOMPATIBLE,
            f"fingerprint del manifiesto {fps} != bundle {bundle.template_fingerprint_id}",
        )

    # el catálogo debe cubrir cada source_metric_id + su cierre transitivo
    missing: set[str] = set()
    stack = [w.source_metric_id for w in bundle.write_instructions]
    seen: set[str] = set()
    coord_to_id = {
        (spec.sheet, spec.cell): mid for mid, spec in bundle.formula_index.items()
    }
    while stack:
        mid = stack.pop()
        if mid in seen:
            continue
        seen.add(mid)
        spec = bundle.formula_index.get(mid)
        if spec is None:
            missing.add(mid)
            continue
        for coord in spec.value_ref_coords:
            dep = coord_to_id.get((spec.sheet, coord))
            if dep is None:
                missing.add(f"{spec.sheet}!{coord}")
            else:
                stack.append(dep)
    if missing:
        raise RuntimeAssetError(
            RUNTIME_ASSET_INVALID,
            f"el catálogo no cubre {len(missing)} métrica(s): {sorted(missing)[:5]}",
        )

    if bundle.zero_write_policy not in _ACCEPTED_ZERO_POLICIES:
        raise RuntimeAssetError(
            RUNTIME_ASSET_INCOMPATIBLE,
            f"zero_write_policy {bundle.zero_write_policy!r} no aceptada",
        )


def _verify_sha256(root: Path, declared: Mapping[str, str]) -> dict[str, str]:
    actual: dict[str, str] = {}
    for name, expected in declared.items():
        path = _require(root / name)
        digest = _sha256(path)
        actual[name] = digest
        if expected and digest != expected:
            raise RuntimeAssetError(
                RUNTIME_ASSET_INVALID,
                f"{name}: sha256 {digest[:12]}… != esperado {str(expected)[:12]}…",
            )
    return actual


def load_runtime_bundle(root: str | Path | None = None) -> RuntimeBundle:
    """Carga y valida el bundle de runtime. Nunca abre el workbook legacy."""
    runtime_root = resolve_runtime_root(root)
    doc = _load_yaml(_require(runtime_root / _BUNDLE_FILE))

    assets = doc.get("assets") or {}
    if not isinstance(assets, dict) or not assets:
        raise RuntimeAssetError(RUNTIME_ASSET_INVALID, "bundle.yaml sin 'assets'")
    asset_sha = _verify_sha256(runtime_root, assets)

    detail = _load_yaml(_require(runtime_root / doc.get("detail_contract", "detail_contract.yaml")))
    columns = detail.get("detail_columns_map") or {}
    criterion_values = {
        (str(entry["sheet"]), str(entry["cell"]).replace("$", "").upper()): entry.get("value")
        for entry in (detail.get("criterion_refs") or [])
    }

    zero_doc = _load_yaml(_require(runtime_root / doc.get("zero_policy", "zero_policy.yaml")))

    manifest = _read_manifest(runtime_root / doc.get("write_manifest", "write_manifest.csv"))
    catalog = _read_catalog(runtime_root / doc.get("metric_catalog", "metric_catalog.csv"))

    bundle = RuntimeBundle(
        version=str(doc.get("version", RUNTIME_DIR_NAME)),
        root=str(runtime_root),
        template_fingerprint_id=str(doc.get("template_fingerprint_id", "")),
        detail_sheet=str(detail.get("detail_sheet", "")),
        detail_columns_map={str(k): str(v) for k, v in columns.items()},
        criterion_values=criterion_values,
        write_instructions=manifest,
        formula_index=catalog,
        expected_instruction_count=int(doc.get("expected_instruction_count", len(manifest))),
        zero_write_policy=str(zero_doc.get("resolution", "UNRESOLVED")),
        estado_filter_status=str(doc.get("estado_filter_status", "PENDING_FUNCTIONAL_CONFIRMATION")),
        asset_sha256=asset_sha,
    )
    if not bundle.detail_sheet or not bundle.detail_columns_map:
        raise RuntimeAssetError(RUNTIME_ASSET_INVALID, "detail_contract.yaml incompleto")
    _validate(bundle)
    return bundle
