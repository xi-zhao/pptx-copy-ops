from __future__ import annotations

from pathlib import Path

from pptx_copy_ops.foreground import (
    CopyReadiness,
    CopyStrategyDecision,
    select_copy_strategy,
)


def _source(**overrides) -> CopyReadiness:
    values = {
        "source_asset_id": "deck.pptx:0",
        "source_pptx": Path("deck.pptx"),
        "source_slide_index": 0,
        "hard_copy_ready": True,
        "foreground_ready": True,
        "source_chrome_ready": False,
        "modifiable_text_slot_count": 2,
        "safe_modification_ops": ("set_shape_text",),
        "recommended_intents": ("exact_reuse", "editable_reuse", "modify_after_copy"),
        "copy_risk_flags": (),
    }
    values.update(overrides)
    return CopyReadiness(**values)


def test_exact_reuse_selects_part_copy_and_discloses_source_background() -> None:
    decision = select_copy_strategy(
        intent="exact_reuse",
        target_template_pptx="target.pptx",
        source=_source(recommended_intents=("exact_reuse",)),
        requires_target_template_inheritance=True,
    )

    assert isinstance(decision, CopyStrategyDecision)
    assert decision.strategy == "exact_part_copy"
    assert decision.preserves_source_background is True
    assert decision.editable_foreground is False
    assert "does_not_inherit_target_template" in decision.risk_flags
    assert "hard_copy_ready" in decision.reasons


def test_editable_reuse_selects_foreground_promote_for_template_inheritance() -> None:
    decision = select_copy_strategy(
        intent="editable_reuse",
        target_template_pptx="target.pptx",
        source=_source(recommended_intents=("editable_reuse",)),
        requires_target_template_inheritance=True,
    )

    assert decision.strategy == "foreground_promote"
    assert decision.preserves_source_background is False
    assert decision.editable_foreground is True
    assert decision.requires_target_template_inheritance is True


def test_source_chrome_reuse_requires_explicit_safe_chrome_capability() -> None:
    blocked = select_copy_strategy(
        intent="source_chrome_reuse",
        target_template_pptx="target.pptx",
        source=_source(source_chrome_ready=False),
    )
    allowed = select_copy_strategy(
        intent="source_chrome_reuse",
        target_template_pptx="target.pptx",
        source=_source(source_chrome_ready=True),
    )

    assert blocked.strategy == "template_synthesize"
    assert blocked.fallback_from == "foreground_with_source_chrome"
    assert "source_chrome_not_ready" in blocked.risk_flags
    assert allowed.strategy == "foreground_with_source_chrome"


def test_modify_after_copy_requires_audited_text_slots_and_allowed_operations() -> None:
    decision = select_copy_strategy(
        intent="modify_after_copy",
        target_template_pptx="target.pptx",
        source=_source(modifiable_text_slot_count=2, safe_modification_ops=("set_shape_text",)),
        required_modification_ops=("set_shape_text",),
        required_text_slot_count=2,
    )
    rejected = select_copy_strategy(
        intent="modify_after_copy",
        target_template_pptx="target.pptx",
        source=_source(modifiable_text_slot_count=1, safe_modification_ops=("set_shape_text",)),
        required_modification_ops=("set_shape_text",),
        required_text_slot_count=2,
    )

    assert decision.strategy == "copy_then_modify"
    assert decision.allowed_modification_ops == ["set_shape_text"]
    assert rejected.strategy == "template_synthesize"
    assert rejected.fallback_from == "copy_then_modify"
    assert "insufficient_modifiable_text_slots" in rejected.risk_flags


def test_missing_or_blocked_source_falls_back_to_template_synthesis() -> None:
    missing = select_copy_strategy(
        intent="editable_reuse",
        target_template_pptx="target.pptx",
        source=None,
    )
    blocked = select_copy_strategy(
        intent="editable_reuse",
        target_template_pptx="target.pptx",
        source=_source(
            hard_copy_ready=False,
            foreground_ready=False,
            source_chrome_ready=False,
            copy_risk_flags=("broken_ooxml",),
        ),
    )

    assert missing.strategy == "template_synthesize"
    assert missing.fallback_from == "editable_reuse"
    assert "missing_source_asset" in missing.risk_flags
    assert blocked.strategy == "template_synthesize"
    assert "foreground_not_ready" in blocked.risk_flags
    assert "broken_ooxml" in blocked.risk_flags
