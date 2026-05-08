from __future__ import annotations

from pathlib import Path

from pptx_copy_ops.copier import SlideCopier, SlideSpec

from .audit import audit_pptx_integrity
from .classifier import classify_element
from .inventory import collect_layer_inventory
from .models import (
    CopyPolicy,
    ForegroundCopyRequest,
    ForegroundCopyResult,
    ForegroundCopyTrace,
)
from .promoter import promote_foreground_elements, remove_slide_background_elements
from .sanitizer import sanitize_pptx_package


def copy_foreground_slide(
    request: ForegroundCopyRequest,
    output_pptx: Path | str,
) -> ForegroundCopyResult:
    source_pptx = Path(request.source_pptx).expanduser().resolve()
    target_pptx = Path(request.target_pptx).expanduser().resolve()
    output = Path(output_pptx).expanduser().resolve()

    inventory = collect_layer_inventory(source_pptx, request.slide_index)
    decisions = [classify_element(element, request.policy) for element in inventory.elements]

    mode = "part" if request.policy.copy_policy == CopyPolicy.EXACT_PART_COPY else "shape"
    copier = SlideCopier(target_template=target_pptx, clear_existing=request.clear_existing)
    copier.copy_slide(SlideSpec(source_pptx, request.slide_index), mode=mode)
    copier.save(output)

    trace = ForegroundCopyTrace(
        decisions=decisions,
        copied_parts=[str(source_pptx)],
    )

    if request.policy.copy_policy != CopyPolicy.EXACT_PART_COPY:
        trace.removed_background_elements.extend(
            remove_slide_background_elements(
                output_pptx=output,
                decisions=decisions,
                target_slide_index=-1,
            )
        )
        promotion = promote_foreground_elements(
            output_pptx=output,
            source_pptx=source_pptx,
            inventory=inventory,
            policy=request.policy,
            target_slide_index=-1,
        )
        trace.promoted_elements.extend(promotion.promoted_elements)
        trace.copied_parts.extend(promotion.copied_parts)

    sanitize_report = sanitize_pptx_package(output)
    trace.sanitized_parts.extend(sanitize_report.sanitized_parts)
    trace.audit_issues.extend(audit_pptx_integrity(output))

    if trace.audit_issues and request.policy.fail_on_audit_error:
        raise ValueError(f"foreground copy audit failed: {trace.audit_issues}")
    return ForegroundCopyResult(output_pptx=output, trace=trace)
