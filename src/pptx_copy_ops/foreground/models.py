from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Literal


class CopyPolicy(str, Enum):
    FOREGROUND_PROMOTE = "foreground_promote"
    EXACT_PART_COPY = "exact_part_copy"
    FOREGROUND_WITH_SOURCE_CHROME = "foreground_with_source_chrome"


CopyStrategy = Literal[
    "exact_part_copy",
    "foreground_promote",
    "foreground_with_source_chrome",
    "copy_then_modify",
    "template_synthesize",
]

CopyIntent = Literal[
    "exact_reuse",
    "editable_reuse",
    "source_chrome_reuse",
    "modify_after_copy",
    "template_synthesize",
]


@dataclass(frozen=True)
class CopyReadiness:
    source_asset_id: str | None = None
    source_pptx: Path | str | None = None
    source_slide_index: int | None = None
    hard_copy_ready: bool = False
    foreground_ready: bool = False
    source_chrome_ready: bool = False
    modifiable_text_slot_count: int = 0
    safe_modification_ops: tuple[str, ...] = ()
    recommended_intents: tuple[CopyIntent | str, ...] = ()
    copy_risk_flags: tuple[str, ...] = ()


@dataclass(frozen=True)
class CopyStrategyDecision:
    strategy: CopyStrategy
    intent: CopyIntent
    source_asset_id: str | None
    source_pptx: str | None
    source_slide_index: int | None
    target_template_pptx: str
    requires_target_template_inheritance: bool
    preserves_source_background: bool
    editable_foreground: bool
    allowed_modification_ops: list[str] = field(default_factory=list)
    risk_flags: list[str] = field(default_factory=list)
    reasons: list[str] = field(default_factory=list)
    fallback_from: str | None = None


class LayerName(str, Enum):
    SLIDE = "slide"
    LAYOUT = "layout"
    MASTER = "master"


class ElementClassification(str, Enum):
    FOREGROUND = "foreground"
    BACKGROUND = "background"
    PLACEHOLDER = "placeholder"
    UNSAFE = "unsafe"


@dataclass(frozen=True)
class ForegroundCopyPolicy:
    copy_policy: CopyPolicy = CopyPolicy.FOREGROUND_PROMOTE
    keep_date_footer: bool = False
    keep_slide_numbers: bool = False
    fail_on_audit_error: bool = True


@dataclass(frozen=True)
class ElementRef:
    layer: LayerName
    part_name: str
    element_index: int
    tag: str
    name: str = ""
    text: str = ""
    extents: tuple[int, int, int, int] | None = None


@dataclass(frozen=True)
class ClassificationDecision:
    element: ElementRef
    classification: ElementClassification
    reason: str


@dataclass
class ForegroundCopyTrace:
    decisions: list[ClassificationDecision] = field(default_factory=list)
    copied_parts: list[str] = field(default_factory=list)
    promoted_elements: list[str] = field(default_factory=list)
    removed_background_elements: list[str] = field(default_factory=list)
    sanitized_parts: list[str] = field(default_factory=list)
    audit_issues: list[str] = field(default_factory=list)


@dataclass
class ForegroundCopyRequest:
    source_pptx: Path
    slide_index: int
    target_pptx: Path
    policy: ForegroundCopyPolicy = field(default_factory=ForegroundCopyPolicy)
    clear_existing: bool = True
    target_background_slide_index: int = 0


@dataclass
class ForegroundCopyResult:
    output_pptx: Path
    trace: ForegroundCopyTrace
