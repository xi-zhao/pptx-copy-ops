"""pptx-copy-ops package."""

from .copier import SlideCopier, SlideCopyMode, SlideSpec, copy_pptx_slides
from .foreground import (
    CopyPolicy,
    ForegroundCopyPolicy,
    ForegroundCopyRequest,
    ForegroundCopyResult,
)

__all__ = [
    "CopyPolicy",
    "ForegroundCopyPolicy",
    "ForegroundCopyRequest",
    "ForegroundCopyResult",
    "SlideCopier",
    "SlideCopyMode",
    "SlideSpec",
    "copy_pptx_slides",
]
