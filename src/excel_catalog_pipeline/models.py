"""Domain models shared by inventory, note, and synchronization code."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Literal

METADATA_FIELDS = (
    "title",
    "subject",
    "author",
    "keywords",
    "categories",
    "comments",
    "sourceFileName",
)
CORE_PROPERTY_BY_METADATA_FIELD = {
    "title": "title",
    "subject": "subject",
    "author": "creator",
    "keywords": "keywords",
    "categories": "category",
    "comments": "description",
}
Direction = Literal[
    "unchanged",
    "pull",
    "push",
    "converged",
    "conflict",
]
FrontmatterTermFormat = Literal["obsidian-link", "plain"]


def _canonical_terms(value: str) -> str:
    normalized = re.sub(r"\s+", " ", value).strip()
    if not normalized:
        return ""
    parts = re.split(r"[;\r\n]+", normalized)
    return "; ".join(dict.fromkeys(part.strip() for part in parts if part.strip()))


@dataclass(frozen=True)
class ReferenceConfig:
    text: str = ""
    files: tuple[str, ...] = ()


@dataclass(frozen=True)
class SourceGenerationConfig:
    generator: str | None = None
    reference: ReferenceConfig = field(default_factory=ReferenceConfig)


@dataclass(frozen=True)
class GeneratorConfig:
    bridge_profile: str = "codex-default"
    prompt_profile: str = "default-ja"
    overrides: dict[str, Any] = field(default_factory=dict)
    reference: ReferenceConfig = field(default_factory=ReferenceConfig)


@dataclass(frozen=True)
class ReferenceInput:
    origin: str
    text: str


@dataclass(frozen=True)
class SourceConfig:
    id: str
    path: Path
    include: tuple[str, ...]
    note_root: Path
    recursive: bool = False
    ignore: tuple[str, ...] = ()
    profile: str = "tkn-obsidian-v1"
    frontmatter_term_format: FrontmatterTermFormat = "obsidian-link"
    rename_adapter: str = "report-only"
    single_workbook: Path | None = None
    single_note: Path | None = None
    generation: SourceGenerationConfig = field(default_factory=SourceGenerationConfig)


@dataclass(frozen=True)
class SyncConfig:
    pull_preserves_user_metadata: bool = True
    delete_missing_notes: bool = False
    delete_missing_workbooks: bool = False
    allow_source_rename: bool = False
    max_extracted_text_chars: int = 12000


@dataclass(frozen=True)
class ContextConfig:
    bridge_profile: str = "codex-default"
    overrides: dict[str, Any] = field(default_factory=dict)
    prompt_profile: str = "default-ja"
    profile_dirs: tuple[str, ...] = ()
    max_images: int = 24
    tile_width_points: int = 1200
    tile_height_points: int = 800
    overlap_points: int = 80
    image_dpi: int = 150
    max_cells: int = 10000
    max_objects: int = 10000
    max_input_chars: int | None = None
    max_workbook_mb: int = 100
    max_reference_chars: int = 20000
    # Resolved invocation data, excluded from persisted configuration.
    generator_id: str | None = None
    reference_inputs: tuple[ReferenceInput, ...] = ()


@dataclass(frozen=True)
class CoverConfig:
    # auto inherits the last successful generation mode; new covers use embedded.
    mode: str = "auto"
    sheet: str | None = None
    range: str | None = None
    width: int | None = None


@dataclass(frozen=True)
class ThumbnailConfig:
    enabled: bool = True
    subdir: str = "thumbnails"
    max_edge: int | None = None
    format: str = "webp"
    quality: int = 85
    lossless: bool = True


@dataclass(frozen=True)
class AppConfig:
    schema_version: str
    sources: tuple[SourceConfig, ...]
    sync: SyncConfig
    loaded_files: tuple[Path, ...] = ()
    context: ContextConfig = field(default_factory=ContextConfig)
    cover: CoverConfig = field(default_factory=CoverConfig)
    thumbnails: ThumbnailConfig = field(default_factory=ThumbnailConfig)
    default_generator: str | None = None
    generators: dict[str, GeneratorConfig] = field(default_factory=dict)
    config_details: dict[str, Any] = field(default_factory=dict, compare=False, repr=False)


def context_settings(config: ContextConfig) -> dict[str, Any]:
    """Only user-configurable shared settings, without invocation data."""
    return {
        key: value
        for key, value in asdict(config).items()
        if key not in {"generator_id", "reference_inputs"}
    }


@dataclass
class SheetInventory:
    """Saved OOXML facts; None means unknown, an empty range means no populated cells."""

    stored_range: str | None = None
    content_range: str | None = None
    populated_cells: int | None = None
    tables: int | None = None
    shapes: int | None = None
    images: int | None = None
    charts: int | None = None
    warnings: list[str] = field(default_factory=list)


@dataclass
class WorkbookInfo:
    path: Path
    relative_path: str
    source_root_id: str
    extension: str
    size_bytes: int
    modified: str
    core: dict[str, str] = field(default_factory=dict)
    custom: dict[str, str] = field(default_factory=dict)
    sheets: list[dict[str, str]] = field(default_factory=list)
    sheet_inventory: dict[str, SheetInventory] = field(default_factory=dict)
    sheet_text: dict[str, list[str]] = field(default_factory=dict)
    sheet_text_status: dict[str, str] = field(default_factory=dict)
    content_fingerprint: str = ""
    read_status: str = "ok"
    warnings: list[str] = field(default_factory=list)

    @property
    def workbook_id(self) -> str:
        return self.custom.get("TknExcelCatalogId", "")

    @property
    def legacy_source_id(self) -> str:
        return f"{self.source_root_id}:{self.relative_path}"

    def metadata(self) -> dict[str, str]:
        return {
            "title": self.core.get("title", ""),
            "subject": self.core.get("subject", ""),
            "author": self.core.get("creator", ""),
            "keywords": _canonical_terms(self.core.get("keywords", "")),
            "categories": _canonical_terms(self.core.get("category", "")),
            "comments": self.core.get("description", ""),
            "sourceFileName": self.relative_path,
        }


@dataclass
class ProxyNote:
    path: Path
    frontmatter: dict[str, Any]
    body: str

    @property
    def note_id(self) -> str:
        return str(self.frontmatter.get("noteId", "")).strip()

    @property
    def source_id(self) -> str:
        return str(self.frontmatter.get("sourceId", "")).strip()


@dataclass(frozen=True)
class FieldDecision:
    field: str
    base: str
    source: str
    note: str
    direction: Direction


@dataclass
class Action:
    status: str
    source_root_id: str
    source_path: str = ""
    note_path: str = ""
    workbook_id: str = ""
    changed_fields: list[str] = field(default_factory=list)
    source_to_note_fields: list[str] = field(default_factory=list)
    note_to_source_fields: list[str] = field(default_factory=list)
    field_differences: list[dict[str, str]] = field(default_factory=list)
    conflict_fields: list[str] = field(default_factory=list)
    message: str = ""
    details: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "sourceRoot": self.source_root_id,
            "sourcePath": self.source_path,
            "notePath": self.note_path,
            "workbookId": self.workbook_id,
            "changedFields": self.changed_fields,
            "sourceToNoteFields": self.source_to_note_fields,
            "noteToSourceFields": self.note_to_source_fields,
            "fieldDifferences": self.field_differences,
            "conflictFields": self.conflict_fields,
            "message": self.message,
            "details": self.details,
        }
