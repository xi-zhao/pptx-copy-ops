# pptx-copy-ops

`pptx-copy-ops` is a small Python library for copying slides across PowerPoint files using native PPTX XML operations (no HTML conversion).

It provides two copy modes:

- `part`: full-slide part-level copy for highest visual fidelity across templates.
- `shape`: blank-slide + shape-level copy fallback.

## Why

Cross-template slide copy is easy to break with:

- dangling relationship IDs
- unregistered slide masters
- `sldMasterId` and `sldLayoutId` collisions

This package wraps those low-level OpenXML steps into a stable API.

## Install

```bash
pip install pptx-copy-ops
```

## Python API

```python
from pptx_copy_ops import SlideCopier, SlideSpec

copier = SlideCopier("/path/to/target-template.pptx", clear_existing=True)
copier.copy_slides(
    [
        SlideSpec("/path/to/theme1.pptx", 0),   # 0-based
        SlideSpec("/path/to/theme2.pptx", 12),  # 0-based
    ],
    mode="part",
)
copier.save("/path/to/output.pptx")
```

## CLI

```bash
python -m pptx_copy_ops.cli \
  --target /path/to/target-template.pptx \
  --source /path/to/theme1.pptx:1 \
  --source /path/to/theme2.pptx:13 \
  --mode part \
  --output /path/to/output.pptx
```

Arguments:

- `--mode`: `part` or `shape`
- `--source`: repeatable, 1-based slide index (`/path/file.pptx:slide_number`)

## Guarantees

- Part-level copy skips `notesSlide` relation to avoid noisy cross-package note dependencies.
- Imported slide masters are auto-registered in `presentation.xml`.
- Layout IDs are normalized globally and kept disjoint from master IDs.

## Development

```bash
cd pptx-copy-ops
pytest -q
python -m build
```

## License

MIT
