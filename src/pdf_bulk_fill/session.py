"""Session configuration persistence (session.json).

Compatible with both the legacy camelCase session.json format
(pdfFileName, templateFileName, sourceFiles) and snake_case keys, while
preserving full replay fidelity (flatten, imageDir, sourceDir, pageIndex,
skipMissingImages, allowTruncation, and exact source file paths).
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
    file_path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"name": self.name, "path": self.path}
        if self.file_path is not None:
            data["filePath"] = self.file_path
        return data


@dataclass
class SessionConfig:
    """Persisted paths and execution options for a bulk PDF fill run."""

    pdf_file: str
    template_file: str
    source_files: list[SourceFileEntry] = field(default_factory=list)
    output_dir: str | None = None
    source_dir: str | None = None
    image_dir: str | None = None
    page_index: int = 0
    flatten: bool = False
    skip_missing_images: bool = False
    allow_truncation: bool = False

    def to_dict(self) -> dict[str, Any]:
        """Serialize to a legacy-compatible JSON dictionary."""
        data: dict[str, Any] = {
            "pdfFileName": self.pdf_file,
            "templateFileName": self.template_file,
            "sourceFiles": [entry.to_dict() for entry in self.source_files],
        }
        if self.output_dir is not None:
            data["outputDir"] = self.output_dir
        if self.source_dir is not None:
            data["sourceDir"] = self.source_dir
        if self.image_dir is not None:
            data["imageDir"] = self.image_dir
        if self.page_index != 0:
            data["pageIndex"] = self.page_index
        if self.flatten:
            data["flatten"] = self.flatten
        if self.skip_missing_images:
            data["skipMissingImages"] = self.skip_missing_images
        if self.allow_truncation:
            data["allowTruncation"] = self.allow_truncation
        return data


def _extract_required_str(raw: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = raw.get(key)
        if isinstance(value, str) and value.strip():
            return value
    raise SessionError(f"Missing or empty required session field: {' / '.join(keys)}")


def _extract_optional_str(raw: dict[str, Any], *keys: str) -> str | None:
    for key in keys:
        if key in raw:
            value = raw[key]
            if value is None:
                return None
            if not isinstance(value, str):
                raise SessionError(f"'{key}' must be a string when provided.")
            return value
    return None


def _extract_optional_bool(raw: dict[str, Any], *keys: str, default: bool = False) -> bool:
    for key in keys:
        if key in raw:
            value = raw[key]
            if not isinstance(value, bool):
                raise SessionError(f"'{key}' must be a boolean when provided.")
            return value
    return default


def _extract_optional_int(raw: dict[str, Any], *keys: str, default: int = 0) -> int:
    for key in keys:
        if key in raw:
            value = raw[key]
            if isinstance(value, bool) or not isinstance(value, int):
                raise SessionError(f"'{key}' must be an integer when provided.")
            return value
    return default


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
        file_path = item.get("filePath", item.get("file_path"))
        if not isinstance(name, str) or not name.strip():
            raise SessionError(f"sourceFiles[{idx}].name must be a non-empty string.")
        if not isinstance(src_path, str) or not src_path.strip():
            raise SessionError(f"sourceFiles[{idx}].path must be a non-empty string.")
        if file_path is not None and (not isinstance(file_path, str) or not file_path.strip()):
            raise SessionError(f"sourceFiles[{idx}].filePath must be a non-empty string.")
        source_files.append(
            SourceFileEntry(name=name, path=src_path, file_path=file_path)
        )

    output_dir = _extract_optional_str(raw, "outputDir", "output_dir")
    source_dir = _extract_optional_str(raw, "sourceDir", "source_dir")
    image_dir = _extract_optional_str(raw, "imageDir", "image_dir")
    page_index = _extract_optional_int(raw, "pageIndex", "page_index", default=0)
    flatten = _extract_optional_bool(raw, "flatten", default=False)
    skip_missing_images = _extract_optional_bool(
        raw, "skipMissingImages", "skip_missing_images", default=False
    )
    allow_truncation = _extract_optional_bool(
        raw, "allowTruncation", "allow_truncation", default=False
    )

    return SessionConfig(
        pdf_file=pdf_file,
        template_file=template_file,
        source_files=source_files,
        output_dir=output_dir,
        source_dir=source_dir,
        image_dir=image_dir,
        page_index=page_index,
        flatten=flatten,
        skip_missing_images=skip_missing_images,
        allow_truncation=allow_truncation,
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
            if entry.file_path:
                return Path(entry.file_path)
            return Path(entry.path) / source_name
    return Path(default_dir) / source_name
