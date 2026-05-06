from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path


class CopyPolicy(str, Enum):
    FOREGROUND_PROMOTE = "foreground_promote"
    EXACT_PART_COPY = "exact_part_copy"
    FOREGROUND_WITH_SOURCE_CHROME = "foreground_with_source_chrome"


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
    sanitized_parts: list[str] = field(default_factory=list)
    audit_issues: list[str] = field(default_factory=list)


@dataclass
class ForegroundCopyRequest:
    source_pptx: Path
    slide_index: int
    target_pptx: Path
    policy: ForegroundCopyPolicy = field(default_factory=ForegroundCopyPolicy)


@dataclass
class ForegroundCopyResult:
    output_pptx: Path
    trace: ForegroundCopyTrace
