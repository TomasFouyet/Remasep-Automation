import json
from dataclasses import asdict, dataclass, field
from datetime import UTC, datetime
from hashlib import sha256
from pathlib import Path


def file_sha256(path: str | Path) -> str:
    digest = sha256()
    with Path(path).open("rb") as file:
        for block in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


@dataclass
class ExecutionAudit:
    period: str
    software_version: str
    template_version: str | None = None
    inputs: list[dict] = field(default_factory=list)
    counters: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    control_status: str | None = None
    created_at_utc: str = field(
        default_factory=lambda: datetime.now(UTC).isoformat()
    )

    def add_input(self, path: str | Path, source_type: str) -> None:
        path = Path(path)
        self.inputs.append(
            {
                "source_type": source_type,
                "filename": path.name,
                "sha256": file_sha256(path),
            }
        )

    def save(self, path: str | Path) -> None:
        Path(path).write_text(
            json.dumps(asdict(self), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
