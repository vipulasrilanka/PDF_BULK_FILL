"""Session configuration persistence (session.json).

Compatible with both the legacy camelCase session.json format
(pdfFileName, templateFileName, sourceFiles) and snake_case keys.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pdf_bulk_fill.errors import SessionError


@dataclass(frozen=True)
class SourceFileEntry:
    """Stored path mapping for a single external source Excel file."""

    name: str
    path: str


@dataclass
class SessionConfig:
    """Persisted paths for a bulk PDF fill run."""

    pdf_file: str
    template_file: str
    source_files: list[SourceFileEntry] = field(default_factory=list)
    output_dir: str | None = None

    def to_dict(self) -> dict[str, Any]:
        """Serialize to a legacy-compatible JSON dictionary."""
        data: dict[str, Any] = {
            "pdfFileName": self.pdf_file,
            "templateFileName": self.template_file,
            "sourceFiles": [
                {"name": entry.name, "path": entry.path}
                for entry in self.source_files
            ],
        }
        if self.output_dir is not None:
            data["outputDir"] = self.output_dir
        return data


def _extract_required_str(raw: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = raw.get(key)
        if isinstance(value, str) and value.strip():
            return value
    raise SessionError(f"Missing or empty required session field: {' / '.join(keys)}")


def load_session(session_path: str | Path) -> SessionConfig:
    """Load and validate a SessionConfig from a JSON file."""
    path = Path(session_path)
    if not path.exists():
        raise SessionError(f"Session file not found: {path}")

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise SessionError(f"Malformed JSON in session file {path}: {exc}") from exc
    except OSError as exc:
        raise SessionError(f"Could not read session file {path}: {exc}") from exc

    if not isinstance(raw, dict):
        raise SessionError(f"Session file {path} must contain a JSON object.")

    pdf_file = _extract_required_str(raw, "pdfFileName", "pdf_file")
    template_file = _extract_required_str(raw, "templateFileName", "template_file")

    raw_sources = raw.get("sourceFiles", raw.get("source_files", []))
    if not isinstance(raw_sources, list):
        raise SessionError("'sourceFiles' must be a list of {name, path} objects.")

    source_files: list[SourceFileEntry] = []
    for idx, item in enumerate(raw_sources):
        if not isinstance(item, dict):
            raise SessionError(f"sourceFiles[{idx}] must be an object.")
        name = item.get("name")
        src_path = item.get("path")
        if not isinstance(name, str) or not name.strip():
            raise SessionError(f"sourceFiles[{idx}].name must be a non-empty string.")
        if not isinstance(src_path, str) or not src_path.strip():
            raise SessionError(f"sourceFiles[{idx}].path must be a non-empty string.")
        source_files.append(SourceFileEntry(name=name, path=src_path))

    output_dir = raw.get("outputDir", raw.get("output_dir"))
    if output_dir is not None and not isinstance(output_dir, str):
        raise SessionError("'outputDir' must be a string when provided.")

    return SessionConfig(
        pdf_file=pdf_file,
        template_file=template_file,
        source_files=source_files,
        output_dir=output_dir,
    )


def save_session(session_path: str | Path, config: SessionConfig) -> Path:
    """Write a SessionConfig to disk as formatted JSON."""
    path = Path(session_path)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(config.to_dict(), indent=4) + "\n", encoding="utf-8")
    except OSError as exc:
        raise SessionError(f"Could not write session file {path}: {exc}") from exc
    return path


def resolve_source_path(
    source_name: str,
    session: SessionConfig,
    default_dir: str | Path,
) -> Path:
    """Return the full path to a source file using session mappings or default_dir."""
    for entry in session.source_files:
        if entry.name == source_name:
            return Path(entry.path) / source_name
    return Path(default_dir) / source_name
