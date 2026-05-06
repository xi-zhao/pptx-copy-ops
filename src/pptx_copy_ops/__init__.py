"""pptx-copy-ops package."""

from .copier import SlideCopier, SlideCopyMode, SlideSpec, copy_pptx_slides
from .foreground import (
    CopyReadiness,
    CopyStrategyDecision,
    CopyPolicy,
    ForegroundCopyPolicy,
    ForegroundCopyRequest,
    ForegroundCopyResult,
    select_copy_strategy,
)

__all__ = [
    "CopyPolicy",
    "CopyReadiness",
    "CopyStrategyDecision",
    "ForegroundCopyPolicy",
    "ForegroundCopyRequest",
    "ForegroundCopyResult",
    "SlideCopier",
    "SlideCopyMode",
    "SlideSpec",
    "copy_pptx_slides",
    "select_copy_strategy",
]
