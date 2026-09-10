"""``FakeWorkbookWriter`` — doble de test del contrato :class:`WorkbookWriter`.

Modela un workbook en memoria (hojas, fórmulas, valores, celdas escribibles,
merges incompatibles, proyecto VBA por módulos). **No abre Excel ni escribe
.xlsm.** Permite probar todo el núcleo de :mod:`remasep.services.excel_writer`
sin Windows.
"""

from __future__ import annotations

import hashlib
from collections.abc import Callable
from contextlib import contextmanager
from dataclasses import dataclass, field

from remasep.services.excel_writer import WorkbookSnapshot
from remasep.services.vba_integrity import (
    EXTRACTION_ABSENT,
    EXTRACTION_EXTRACTED,
    VbaModule,
    VbaProject,
)

_DEFAULT_VBA = {"Módulo1": "Sub PROTEGER()\nEnd Sub", "ThisWorkbook": ""}


def _sha(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


@dataclass
class FakeWorkbookModel:
    sheet_names: list[str] = field(default_factory=lambda: ["REMASEP_OD", "CONTROL"])
    formula_map: dict[tuple[str, str], str] = field(default_factory=dict)
    values: dict[tuple[str, str], object] = field(default_factory=dict)
    writable_cells: set[tuple[str, str]] = field(default_factory=set)
    incompatible_merges: set[tuple[str, str]] = field(default_factory=set)
    vba_present: bool = True
    vba_payload_sha256: str | None = "fake-vba-sha"
    # código VBA por módulo (name -> cuerpo de código); None -> conjunto por defecto.
    vba_modules: dict[str, str] | None = None
    # si es False, el proyecto VBA se reporta como no extraíble semánticamente.
    vba_semantic_available: bool = True
    structural_fingerprint_id: str = "stf:dc624775927d4d4d"
    # comportamiento simulado del recálculo: (sheet, cell) -> callable(values) -> value
    recalculators: dict[tuple[str, str], object] = field(default_factory=dict)
    # inyección de fallos: se aplica al modelo dentro de save() (simula que Excel
    # alteró una fórmula, perdió el VBA, cambió el fingerprint, etc.)
    mutate_after_save: Callable[[FakeWorkbookModel], None] | None = None

    def copy(self) -> FakeWorkbookModel:
        return FakeWorkbookModel(
            sheet_names=list(self.sheet_names),
            formula_map=dict(self.formula_map),
            values=dict(self.values),
            writable_cells=set(self.writable_cells),
            incompatible_merges=set(self.incompatible_merges),
            vba_present=self.vba_present,
            vba_payload_sha256=self.vba_payload_sha256,
            vba_modules=None if self.vba_modules is None else dict(self.vba_modules),
            vba_semantic_available=self.vba_semantic_available,
            structural_fingerprint_id=self.structural_fingerprint_id,
            recalculators=dict(self.recalculators),
            mutate_after_save=self.mutate_after_save,
        )

    def vba_project(self) -> VbaProject:
        if not self.vba_present:
            return VbaProject(
                present=False, payload_sha256=None, modules=(),
                extraction_status=EXTRACTION_ABSENT,
            )
        if not self.vba_semantic_available:
            return VbaProject(
                present=True, payload_sha256=self.vba_payload_sha256, modules=(),
                extraction_status="UNAVAILABLE_PARSE_ERROR",
                note="fake: extracción semántica deshabilitada",
            )
        src = _DEFAULT_VBA if self.vba_modules is None else self.vba_modules
        modules = tuple(
            VbaModule(name=name, code_sha256=_sha(code), attributes_sha256=_sha(""))
            for name, code in sorted(src.items())
        )
        return VbaProject(
            present=True,
            payload_sha256=self.vba_payload_sha256,
            modules=modules,
            extraction_status=EXTRACTION_EXTRACTED,
        )


class FakeWorkbookWriter:
    """Implementación en memoria del contrato ``WorkbookWriter``."""

    def __init__(self, model: FakeWorkbookModel, path: str) -> None:
        self._model = model.copy()
        self._path = path
        self.saved = False
        self.recalculated = False
        self.closed = False
        self.writes: list[tuple[str, str, object]] = []

    # -- lectura --------------------------------------------------------
    def sheet_names(self) -> list[str]:
        return list(self._model.sheet_names)

    def has_sheet(self, sheet: str) -> bool:
        return sheet in self._model.sheet_names

    def cell_has_formula(self, sheet: str, cell: str) -> bool:
        return (sheet, cell) in self._model.formula_map

    def cell_in_incompatible_merge(self, sheet: str, cell: str) -> bool:
        return (sheet, cell) in self._model.incompatible_merges

    def cell_is_writable(self, sheet: str, cell: str) -> bool:
        return (sheet, cell) in self._model.writable_cells

    def read_cell(self, sheet: str, cell: str) -> object:
        return self._model.values.get((sheet, cell))

    # -- escritura -----------------------------------------------------
    def write_value2(self, sheet: str, cell: str, value: object) -> None:
        if self.closed:
            raise RuntimeError("writer cerrado")
        if (sheet, cell) in self._model.formula_map:
            raise RuntimeError(f"no se sobrescribe fórmula: {sheet}!{cell}")
        self._model.values[(sheet, cell)] = value
        self.writes.append((sheet, cell, value))

    def recalculate(self) -> None:
        self.recalculated = True
        for key, fn in self._model.recalculators.items():
            self._model.values[key] = fn(self._model.values)

    def save(self) -> None:
        if not self.recalculated:
            raise RuntimeError("save() antes de recalculate()")
        if self._model.mutate_after_save is not None:
            self._model.mutate_after_save(self._model)
        self.saved = True

    def snapshot(self) -> WorkbookSnapshot:
        project = self._model.vba_project()
        return WorkbookSnapshot(
            path=self._path,
            sheet_names=tuple(self._model.sheet_names),
            formula_map=dict(self._model.formula_map),
            values=dict(self._model.values),
            vba_present=project.present,
            vba_payload_sha256=project.payload_sha256,
            structural_fingerprint_id=self._model.structural_fingerprint_id,
            file_sha256=None,
            vba_project=project,
        )

    def close(self) -> None:
        self.closed = True


class FakeWriterHarness:
    """Provee ``open_writer`` e ``inspect`` para :func:`excel_writer.generate`.

    Recuerda el snapshot post-``save`` del último writer y lo devuelve tanto para
    la copia de trabajo como (tras la promoción atómica) para la salida final.
    El snapshot de la plantilla es el estado inicial del modelo.
    """

    def __init__(self, model: FakeWorkbookModel, *, fail_during_write: bool = False) -> None:
        self.model = model
        self.fail_during_write = fail_during_write
        self.last_writer: FakeWorkbookWriter | None = None
        self._template_snapshot = FakeWorkbookWriter(model, "<template>").snapshot()

    @contextmanager
    def open_writer(self, path):
        writer = FakeWorkbookWriter(self.model, str(path))
        self.last_writer = writer
        try:
            if self.fail_during_write:
                raise RuntimeError("fallo simulado durante la escritura")
            yield writer
        finally:
            writer.close()

    def inspect(self, path) -> WorkbookSnapshot:
        if self.last_writer is not None and self.last_writer.saved:
            return self.last_writer.snapshot()
        return self._template_snapshot
