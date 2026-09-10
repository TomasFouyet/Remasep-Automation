"""Integridad **semántica** del proyecto VBA (Sprint 3.7B — patch final).

Comparar el SHA256 bruto de ``xl/vbaProject.bin`` es un criterio incorrecto: un
``Save`` legítimo de Excel regenera los streams de caché compilada
(``__SRP_0..n``, ``_VBA_PROJECT``, la P-code de cada módulo) **sin** que cambie
una línea de código de macro. Este módulo compara, en cambio, el **código
fuente descomprimido por módulo**:

- se abre ``vbaProject.bin`` (OLE Compound File) con ``olefile`` — sólo lectura;
- se descomprime el *CompressedContainer* de cada stream ``VBA/<módulo>`` según
  MS-OVBA §2.4.1 (sin ejecutar nada, sin COM, sin VBIDE, sin tocar el Trust
  Center ni "Trust access to the VBA project object model");
- se separa el bloque de atributos (``Attribute VB_*``) del cuerpo de código y se
  hashea cada parte por separado.

La integridad del writer **falla** si: el proyecto VBA existía y desaparece (o
aparece), cambia el conjunto de módulos, o cambia el **código** de un módulo. Un
cambio sólo del binario ``vbaProject.bin`` (o un reordenamiento de atributos /
controles) queda como **diagnóstico/warning** — pero **sólo** cuando la
comparación semántica de código sí se pudo ejecutar y demostró equivalencia.

**Fail-closed**: si el proyecto VBA existe pero **no** se pudo extraer/comparar
el código (sin ``olefile``, o parseo fallido) el resultado es
``VBA_INTEGRITY_FAIL`` con ``reason = VBA_SEMANTIC_CHECK_UNAVAILABLE``. No se
promueve el workbook al output final; se conservan presencia y hashes como
evidencia. **Nunca** se asume equivalencia sólo porque ``vbaProject.bin`` siga
presente.
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from remasep.core.errors import RemasepError

try:  # dependencia lean, pura Python (~114 KB)
    import olefile  # type: ignore[import-untyped]

    _OLEFILE_AVAILABLE = True
except ImportError:  # pragma: no cover - defensivo
    olefile = None  # type: ignore[assignment]
    _OLEFILE_AVAILABLE = False

_VBA_ENTRY = "xl/vbaproject.bin"

# estados de extracción de un proyecto
EXTRACTION_EXTRACTED = "EXTRACTED"
EXTRACTION_ABSENT = "ABSENT"
EXTRACTION_UNAVAILABLE_NO_OLEFILE = "UNAVAILABLE_NO_OLEFILE"
EXTRACTION_UNAVAILABLE_PARSE_ERROR = "UNAVAILABLE_PARSE_ERROR"

# estados de la comparación (``VbaComparison.status``): sólo PASS o FAIL.
VBA_PASS = "VBA_INTEGRITY_PASS"
VBA_FAIL = "VBA_INTEGRITY_FAIL"
# token de motivo cuando el FAIL se debe a que no se pudo verificar el código
# (fail-closed); nunca es un ``status`` por sí solo.
VBA_SEMANTIC_UNAVAILABLE = "VBA_SEMANTIC_CHECK_UNAVAILABLE"


class VbaDecompressionError(RemasepError):
    """El *CompressedContainer* MS-OVBA está corrupto o truncado."""


def olefile_available() -> bool:
    return _OLEFILE_AVAILABLE


# ---------------------------------------------------------------------------
# MS-OVBA §2.4.1 — descompresión del CompressedContainer
# ---------------------------------------------------------------------------


def decompress_vba_container(data: bytes) -> bytes:
    """Descomprime un *CompressedContainer* (MS-OVBA §2.4.1). Solo lectura."""
    if len(data) < 1 or data[0] != 0x01:
        raise VbaDecompressionError("falta la firma 0x01 del CompressedContainer")
    out = bytearray()
    i = 1
    n = len(data)
    while i < n:
        if i + 2 > n:
            raise VbaDecompressionError("cabecera de chunk truncada")
        header = int.from_bytes(data[i : i + 2], "little")
        i += 2
        chunk_size = (header & 0x0FFF) + 3
        compressed_flag = (header & 0x8000) != 0
        if (header >> 12) & 0x7 != 0b011:
            raise VbaDecompressionError("firma de CompressedChunkHeader inválida")
        data_len = chunk_size - 2
        chunk = data[i : i + data_len]
        i += data_len
        if not compressed_flag:
            out.extend(chunk[:4096])
            continue
        chunk_start = len(out)
        j = 0
        while j < len(chunk):
            flags = chunk[j]
            j += 1
            for bit in range(8):
                if j >= len(chunk):
                    break
                if not (flags >> bit) & 1:
                    out.append(chunk[j])
                    j += 1
                else:
                    if j + 2 > len(chunk):
                        raise VbaDecompressionError("CopyToken truncado")
                    token = int.from_bytes(chunk[j : j + 2], "little")
                    j += 2
                    difference = len(out) - chunk_start
                    bit_count = max((difference - 1).bit_length(), 4)
                    length_mask = 0xFFFF >> bit_count
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    src = len(out) - offset
                    if src < 0:
                        raise VbaDecompressionError("offset de CopyToken fuera de rango")
                    for k in range(length):
                        out.append(out[src + k])
    return bytes(out)


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class VbaModule:
    name: str
    code_sha256: str
    attributes_sha256: str
    attribute_lines: tuple[str, ...] = ()


@dataclass(frozen=True)
class VbaProject:
    present: bool
    payload_sha256: str | None
    modules: tuple[VbaModule, ...]
    extraction_status: str
    note: str = ""

    @property
    def semantic_available(self) -> bool:
        return self.extraction_status == EXTRACTION_EXTRACTED

    @property
    def module_names(self) -> tuple[str, ...]:
        return tuple(m.name for m in self.modules)

    def as_dict(self) -> dict:
        return {
            "present": self.present,
            "payload_sha256": self.payload_sha256,
            "extraction_status": self.extraction_status,
            "module_count": len(self.modules),
            "module_names": list(self.module_names),
            "note": self.note,
        }


def _absent_project() -> VbaProject:
    return VbaProject(
        present=False, payload_sha256=None, modules=(), extraction_status=EXTRACTION_ABSENT
    )


# ---------------------------------------------------------------------------
# Lectura del proyecto VBA
# ---------------------------------------------------------------------------


def _read_payload_bytes(path: Path) -> bytes | None:
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as archive:
            by_lower = {name.lower(): name for name in archive.namelist()}
            key = by_lower.get(_VBA_ENTRY)
            return archive.read(key) if key else None
    if path.is_file():
        return path.read_bytes()
    return None


def _normalize_lines(raw: bytes) -> list[bytes]:
    unified = raw.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    return [line.rstrip() for line in unified.split(b"\n")]


def _split_module(source: bytes) -> tuple[tuple[bytes, ...], bytes]:
    lines = _normalize_lines(source)
    k = 0
    while k < len(lines) and lines[k].startswith(b"Attribute "):
        k += 1
    attributes = tuple(sorted(lines[:k]))
    body = b"\n".join(lines[k:]).strip(b"\n")
    return attributes, body


def _decode_name(attributes: tuple[bytes, ...]) -> str | None:
    for line in attributes:
        if line.startswith(b'Attribute VB_Name = "') and line.endswith(b'"'):
            return line[len(b'Attribute VB_Name = "') : -1].decode("cp1252", "replace")
    return None


def _module_source(stream: bytes) -> bytes | None:
    """Localiza y descomprime el código fuente dentro de un stream ``VBA/<módulo>``.

    El stream es ``[PerformanceCache][CompressedContainer]``; el contenedor
    empieza por ``0x01`` y descomprime a texto que empieza por ``Attribute VB_``.
    """
    for offset in range(len(stream)):
        if stream[offset] != 0x01:
            continue
        try:
            decoded = decompress_vba_container(stream[offset:])
        except VbaDecompressionError:
            continue
        if decoded[:13] == b"Attribute VB_":
            return decoded
    return None


def _extract_modules(payload: bytes) -> list[VbaModule]:
    modules: list[VbaModule] = []
    ole = olefile.OleFileIO(io.BytesIO(payload))
    try:
        for entry in ole.listdir(streams=True, storages=False):
            if len(entry) < 2 or entry[0] != "VBA":
                continue
            leaf = entry[-1]
            if leaf in ("dir", "_VBA_PROJECT") or leaf.upper().startswith("__SRP_"):
                continue
            source = _module_source(ole.openstream(entry).read())
            if source is None:
                continue
            attributes, body = _split_module(source)
            modules.append(
                VbaModule(
                    name=_decode_name(attributes) or leaf,
                    code_sha256=hashlib.sha256(body).hexdigest(),
                    attributes_sha256=hashlib.sha256(b"\n".join(attributes)).hexdigest(),
                    attribute_lines=tuple(a.decode("cp1252", "replace") for a in attributes),
                )
            )
    finally:
        ole.close()
    modules.sort(key=lambda m: m.name)
    return modules


def read_vba_project(source: str | Path) -> VbaProject:
    """Lee el proyecto VBA de un ``.xlsm`` (o de un ``vbaProject.bin`` suelto)."""
    path = Path(source)
    payload = _read_payload_bytes(path)
    if payload is None:
        return _absent_project()
    payload_sha = hashlib.sha256(payload).hexdigest()
    if not _OLEFILE_AVAILABLE:
        return VbaProject(
            present=True,
            payload_sha256=payload_sha,
            modules=(),
            extraction_status=EXTRACTION_UNAVAILABLE_NO_OLEFILE,
            note="olefile no está instalado; no hay comparación semántica de VBA",
        )
    try:
        modules = _extract_modules(payload)
    except Exception as exc:  # noqa: BLE001 - se resume a un estado, no traceback
        return VbaProject(
            present=True,
            payload_sha256=payload_sha,
            modules=(),
            extraction_status=EXTRACTION_UNAVAILABLE_PARSE_ERROR,
            note=f"{exc.__class__.__name__}: {exc}",
        )
    return VbaProject(
        present=True,
        payload_sha256=payload_sha,
        modules=tuple(modules),
        extraction_status=EXTRACTION_EXTRACTED,
    )


# ---------------------------------------------------------------------------
# Comparación
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class VbaComparison:
    present_before: bool
    present_after: bool
    payload_sha256_before: str | None
    payload_sha256_after: str | None
    payload_stable: bool
    semantic_available: bool
    module_names_before: tuple[str, ...]
    module_names_after: tuple[str, ...]
    added_modules: tuple[str, ...]
    removed_modules: tuple[str, ...]
    code_changed_modules: tuple[str, ...]
    attributes_changed_modules: tuple[str, ...]
    status: str
    fail_reasons: tuple[str, ...] = ()
    warnings: tuple[str, ...] = field(default_factory=tuple)

    @property
    def ok(self) -> bool:
        """Sólo ``VBA_INTEGRITY_PASS`` pasa. La comparación semántica no
        disponible es **fail-closed** (``VBA_INTEGRITY_FAIL``)."""
        return self.status == VBA_PASS

    def as_dict(self) -> dict:
        return {
            "status": self.status,
            "present_before": self.present_before,
            "present_after": self.present_after,
            "vba_binary_payload_sha_before": self.payload_sha256_before,
            "vba_binary_payload_sha_after": self.payload_sha256_after,
            "vba_binary_payload_stable": self.payload_stable,
            "semantic_available": self.semantic_available,
            "module_names_before": list(self.module_names_before),
            "module_names_after": list(self.module_names_after),
            "added_modules": list(self.added_modules),
            "removed_modules": list(self.removed_modules),
            "code_changed_modules": list(self.code_changed_modules),
            "attributes_changed_modules": list(self.attributes_changed_modules),
            "fail_reasons": list(self.fail_reasons),
            "warnings": list(self.warnings),
        }


def compare_vba_projects(
    before: VbaProject | None, after: VbaProject | None
) -> VbaComparison:
    before = before or _absent_project()
    after = after or _absent_project()

    payload_stable = before.payload_sha256 == after.payload_sha256
    fail: list[str] = []
    warn: list[str] = []

    if before.present and not after.present:
        fail.append("VBA_LOST: el proyecto VBA existía y desapareció")
    if after.present and not before.present:
        fail.append("VBA_ADDED: apareció un proyecto VBA que no estaba antes")

    added: tuple[str, ...] = ()
    removed: tuple[str, ...] = ()
    code_changed: tuple[str, ...] = ()
    attrs_changed: tuple[str, ...] = ()

    semantic = (
        before.present
        and after.present
        and before.semantic_available
        and after.semantic_available
    )

    if semantic:
        bmap = {m.name: m for m in before.modules}
        amap = {m.name: m for m in after.modules}
        added = tuple(sorted(set(amap) - set(bmap)))
        removed = tuple(sorted(set(bmap) - set(amap)))
        common = sorted(set(bmap) & set(amap))
        code_changed = tuple(m for m in common if bmap[m].code_sha256 != amap[m].code_sha256)
        attrs_changed = tuple(
            m for m in common if bmap[m].attributes_sha256 != amap[m].attributes_sha256
        )
        if added or removed:
            fail.append(f"VBA_MODULE_SET_CHANGED: +{list(added)} -{list(removed)}")
        if code_changed:
            fail.append(f"VBA_MODULE_SOURCE_CHANGED: {list(code_changed)}")
        if not fail:
            if attrs_changed:
                warn.append(
                    f"VBA_MODULE_ATTRIBUTES_CHANGED: {list(attrs_changed)} "
                    "(reordenamiento de atributos/controles; el código no cambió)"
                )
            if not payload_stable:
                warn.append(
                    "VBA_BINARY_PAYLOAD_CHANGED_SEMANTIC_EQUIVALENT: el Save de Excel "
                    "regenera __SRP_*/_VBA_PROJECT; el código VBA es equivalente"
                )
        status = VBA_FAIL if fail else VBA_PASS
    elif before.present and after.present:
        # fail-closed: el proyecto VBA existe pero no se pudo extraer/comparar el
        # CÓDIGO. No se asume equivalencia por el mero hecho de que
        # vbaProject.bin siga presente: se bloquea la promoción al output final
        # y se conservan presencia + hashes binarios como evidencia.
        fail.append(
            f"{VBA_SEMANTIC_UNAVAILABLE}: el proyecto VBA existe pero no se pudo "
            "extraer ni comparar su código "
            f"(before={before.extraction_status}, after={after.extraction_status}). "
            "Instala `olefile` o revisá el workbook. No se genera la salida final. "
            f"Evidencia: presente_antes={before.present}, presente_después={after.present}, "
            f"vba_binary_payload_sha_before={before.payload_sha256}, "
            f"vba_binary_payload_sha_after={after.payload_sha256}."
        )
        if not payload_stable:
            warn.append("VBA_BINARY_PAYLOAD_CHANGED (diagnóstico; sin comparación semántica)")
        status = VBA_FAIL
    else:
        status = VBA_FAIL if fail else VBA_PASS

    return VbaComparison(
        present_before=before.present,
        present_after=after.present,
        payload_sha256_before=before.payload_sha256,
        payload_sha256_after=after.payload_sha256,
        payload_stable=payload_stable,
        semantic_available=semantic,
        module_names_before=before.module_names,
        module_names_after=after.module_names,
        added_modules=added,
        removed_modules=removed,
        code_changed_modules=code_changed,
        attributes_changed_modules=attrs_changed,
        status=status,
        fail_reasons=tuple(fail),
        warnings=tuple(warn),
    )
