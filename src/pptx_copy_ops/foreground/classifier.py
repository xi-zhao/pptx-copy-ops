from __future__ import annotations

from .models import (
    ClassificationDecision,
    CopyPolicy,
    ElementClassification,
    ElementRef,
    ForegroundCopyPolicy,
    LayerName,
)

SLIDE_WIDTH_EMU = 12_192_000
SLIDE_HEIGHT_EMU = 6_858_000


def _is_placeholder(element: ElementRef, policy: ForegroundCopyPolicy) -> bool:
    name = element.name.lower()
    text = element.text.strip()

    if text in {"‹#›", "<#>", "#"}:
        return not policy.keep_slide_numbers
    if ("date" in name or "日期" in element.name or "footer" in name or "页脚" in element.name):
        return not policy.keep_date_footer
    if policy.copy_policy == CopyPolicy.EXACT_PART_COPY:
        return False
    return (
        "placeholder" in name
        or "占位符" in element.name
        or "单击此处编辑母版" in text
    )


def _is_large_canvas(element: ElementRef) -> bool:
    if element.extents is None:
        return False
    x, y, cx, cy = element.extents
    nearly_full_width = abs(x) < 200_000 and cx > int(SLIDE_WIDTH_EMU * 0.85)
    nearly_full_height = abs(y) < 200_000 and cy > int(SLIDE_HEIGHT_EMU * 0.75)
    large_band = nearly_full_width and cy > 500_000
    return large_band or (nearly_full_width and nearly_full_height)


def classify_element(
    element: ElementRef,
    policy: ForegroundCopyPolicy,
) -> ClassificationDecision:
    if element.layer == LayerName.SLIDE:
        return ClassificationDecision(element, ElementClassification.FOREGROUND, "slide-local content")

    if _is_placeholder(element, policy):
        return ClassificationDecision(element, ElementClassification.PLACEHOLDER, "master/layout placeholder")

    if element.tag == "pic":
        if policy.copy_policy == CopyPolicy.FOREGROUND_PROMOTE and _is_large_canvas(element):
            return ClassificationDecision(element, ElementClassification.BACKGROUND, "large background picture")
        return ClassificationDecision(element, ElementClassification.FOREGROUND, "picture foreground")

    if element.text.strip():
        return ClassificationDecision(element, ElementClassification.FOREGROUND, "text foreground")

    if policy.copy_policy == CopyPolicy.FOREGROUND_PROMOTE and _is_large_canvas(element):
        return ClassificationDecision(element, ElementClassification.BACKGROUND, "large no-text canvas")

    return ClassificationDecision(element, ElementClassification.FOREGROUND, "small non-placeholder shape")
