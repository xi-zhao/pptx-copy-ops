from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable, Mapping

from .models import CopyIntent, CopyReadiness, CopyStrategy, CopyStrategyDecision

BLOCKING_COPY_RISKS = {
    "broken_ooxml",
    "copy_blocked",
    "missing_source_part",
    "audit_failed",
}


def _read(source: CopyReadiness | Mapping[str, Any], name: str, default: Any = None) -> Any:
    if isinstance(source, Mapping):
        return source.get(name, default)
    return getattr(source, name, default)


def _source_text(source: CopyReadiness | Mapping[str, Any], name: str) -> str | None:
    value = _read(source, name)
    if value is None:
        return None
    return str(value)


def _source_int(source: CopyReadiness | Mapping[str, Any], name: str) -> int | None:
    value = _read(source, name)
    if value is None:
        return None
    return int(value)


def _source_tuple(source: CopyReadiness | Mapping[str, Any], name: str) -> tuple[str, ...]:
    value = _read(source, name, ())
    if value is None:
        return ()
    return tuple(str(item) for item in value)


def _decision(
    *,
    strategy: CopyStrategy,
    intent: CopyIntent,
    target_template_pptx: str | Path,
    source: CopyReadiness | Mapping[str, Any] | None,
    requires_target_template_inheritance: bool,
    preserves_source_background: bool,
    editable_foreground: bool,
    allowed_modification_ops: Iterable[str] = (),
    risk_flags: Iterable[str] = (),
    reasons: Iterable[str] = (),
    fallback_from: str | None = None,
) -> CopyStrategyDecision:
    return CopyStrategyDecision(
        strategy=strategy,
        intent=intent,
        source_asset_id=_source_text(source, "source_asset_id") if source is not None else None,
        source_pptx=_source_text(source, "source_pptx") if source is not None else None,
        source_slide_index=_source_int(source, "source_slide_index") if source is not None else None,
        target_template_pptx=str(target_template_pptx),
        requires_target_template_inheritance=requires_target_template_inheritance,
        preserves_source_background=preserves_source_background,
        editable_foreground=editable_foreground,
        allowed_modification_ops=list(allowed_modification_ops),
        risk_flags=list(dict.fromkeys(str(flag) for flag in risk_flags)),
        reasons=list(dict.fromkeys(str(reason) for reason in reasons)),
        fallback_from=fallback_from,
    )


def _fallback(
    *,
    intent: CopyIntent,
    target_template_pptx: str | Path,
    source: CopyReadiness | Mapping[str, Any] | None,
    requires_target_template_inheritance: bool,
    risk_flags: Iterable[str],
    reasons: Iterable[str] = (),
    fallback_from: str | None = None,
) -> CopyStrategyDecision:
    return _decision(
        strategy="template_synthesize",
        intent=intent,
        target_template_pptx=target_template_pptx,
        source=source,
        requires_target_template_inheritance=requires_target_template_inheritance,
        preserves_source_background=False,
        editable_foreground=True,
        risk_flags=risk_flags,
        reasons=reasons,
        fallback_from=fallback_from or intent,
    )


def select_copy_strategy(
    *,
    intent: CopyIntent,
    target_template_pptx: str | Path,
    source: CopyReadiness | Mapping[str, Any] | None,
    requires_target_template_inheritance: bool = False,
    required_modification_ops: Iterable[str] = (),
    required_text_slot_count: int = 0,
) -> CopyStrategyDecision:
    """Choose an explicit product-level copy strategy for one slide."""

    if intent == "template_synthesize":
        return _decision(
            strategy="template_synthesize",
            intent=intent,
            target_template_pptx=target_template_pptx,
            source=source,
            requires_target_template_inheritance=requires_target_template_inheritance,
            preserves_source_background=False,
            editable_foreground=True,
            reasons=["template_synthesis_requested"],
        )

    if source is None:
        return _fallback(
            intent=intent,
            target_template_pptx=target_template_pptx,
            source=None,
            requires_target_template_inheritance=requires_target_template_inheritance,
            risk_flags=["missing_source_asset"],
            reasons=["no_source_asset"],
        )

    source_risks = list(_source_tuple(source, "copy_risk_flags"))

    if intent == "exact_reuse":
        if bool(_read(source, "hard_copy_ready", False)):
            risks = source_risks
            if requires_target_template_inheritance:
                risks.append("does_not_inherit_target_template")
            return _decision(
                strategy="exact_part_copy",
                intent=intent,
                target_template_pptx=target_template_pptx,
                source=source,
                requires_target_template_inheritance=requires_target_template_inheritance,
                preserves_source_background=True,
                editable_foreground=False,
                risk_flags=risks,
                reasons=["hard_copy_ready"],
            )
        return _fallback(
            intent=intent,
            target_template_pptx=target_template_pptx,
            source=source,
            requires_target_template_inheritance=requires_target_template_inheritance,
            risk_flags=[*source_risks, "hard_copy_not_ready"],
            reasons=["hard_copy_unavailable"],
            fallback_from="exact_part_copy",
        )

    if intent == "editable_reuse":
        if bool(_read(source, "foreground_ready", False)):
            return _decision(
                strategy="foreground_promote",
                intent=intent,
                target_template_pptx=target_template_pptx,
                source=source,
                requires_target_template_inheritance=requires_target_template_inheritance,
                preserves_source_background=False,
                editable_foreground=True,
                risk_flags=source_risks,
                reasons=["foreground_ready"],
            )
        return _fallback(
            intent=intent,
            target_template_pptx=target_template_pptx,
            source=source,
            requires_target_template_inheritance=requires_target_template_inheritance,
            risk_flags=[*source_risks, "foreground_not_ready"],
            reasons=["foreground_unavailable"],
            fallback_from="foreground_promote",
        )

    if intent == "source_chrome_reuse":
        if bool(_read(source, "source_chrome_ready", False)):
            return _decision(
                strategy="foreground_with_source_chrome",
                intent=intent,
                target_template_pptx=target_template_pptx,
                source=source,
                requires_target_template_inheritance=requires_target_template_inheritance,
                preserves_source_background=False,
                editable_foreground=True,
                risk_flags=source_risks,
                reasons=["source_chrome_ready"],
            )
        return _fallback(
            intent=intent,
            target_template_pptx=target_template_pptx,
            source=source,
            requires_target_template_inheritance=requires_target_template_inheritance,
            risk_flags=[*source_risks, "source_chrome_not_ready"],
            reasons=["source_chrome_unavailable"],
            fallback_from="foreground_with_source_chrome",
        )

    if intent == "modify_after_copy":
        allowed_ops = set(_source_tuple(source, "safe_modification_ops"))
        required_ops = tuple(str(op) for op in required_modification_ops)
        missing_ops = [op for op in required_ops if op not in allowed_ops]
        modifiable_count = int(_read(source, "modifiable_text_slot_count", 0) or 0)
        risks = list(source_risks)
        blocking_source_risks = [
            risk for risk in source_risks if risk in BLOCKING_COPY_RISKS
        ]
        if missing_ops:
            risks.append("unsafe_modification_ops")
        if modifiable_count < required_text_slot_count:
            risks.append("insufficient_modifiable_text_slots")
        if missing_ops or modifiable_count < required_text_slot_count or blocking_source_risks:
            return _fallback(
                intent=intent,
                target_template_pptx=target_template_pptx,
                source=source,
                requires_target_template_inheritance=requires_target_template_inheritance,
                risk_flags=risks,
                reasons=["copy_then_modify_unavailable"],
                fallback_from="copy_then_modify",
            )
        return _decision(
            strategy="copy_then_modify",
            intent=intent,
            target_template_pptx=target_template_pptx,
            source=source,
            requires_target_template_inheritance=requires_target_template_inheritance,
            preserves_source_background=not bool(_read(source, "foreground_ready", False)),
            editable_foreground=bool(_read(source, "foreground_ready", False)),
            allowed_modification_ops=required_ops,
            risk_flags=risks,
            reasons=["audited_modification_slots_ready", "safe_modification_ops_ready"],
        )

    return _fallback(
        intent=intent,
        target_template_pptx=target_template_pptx,
        source=source,
        requires_target_template_inheritance=requires_target_template_inheritance,
        risk_flags=[*source_risks, "unknown_copy_intent"],
        reasons=["unknown_copy_intent"],
    )
