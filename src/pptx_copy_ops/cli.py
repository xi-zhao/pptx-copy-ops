from __future__ import annotations

import argparse
from pathlib import Path

from .copier import SlideSpec, copy_pptx_slides


def _parse_source_spec(raw: str) -> SlideSpec:
    if ":" not in raw:
        raise ValueError(f"Invalid --source spec: {raw}. Expected '/path/file.pptx:slide_number'")
    path_part, index_part = raw.rsplit(":", 1)
    try:
        one_based_index = int(index_part)
    except ValueError as exc:
        raise ValueError(f"Invalid slide number in --source spec: {raw}") from exc
    if one_based_index <= 0:
        raise ValueError(f"Slide number must be >= 1 in --source spec: {raw}")
    return SlideSpec(source_path=Path(path_part), slide_index=one_based_index - 1)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Copy PPTX slides across templates.")
    parser.add_argument("--target", required=True, help="Target template PPTX path")
    parser.add_argument(
        "--source",
        action="append",
        required=True,
        help="Source slide spec '/path/to/source.pptx:slide_number(1-based)'",
    )
    parser.add_argument(
        "--mode",
        choices=["shape", "part"],
        default="part",
        help="Copy mode: shape (safe fallback) or part (full slide copy)",
    )
    parser.add_argument("--output", required=True, help="Output PPTX path")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    specs = [_parse_source_spec(raw) for raw in args.source]
    destination = copy_pptx_slides(
        target_template=args.target,
        sources=specs,
        output_pptx=args.output,
        mode=args.mode,
        clear_existing=True,
    )
    print(destination)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
