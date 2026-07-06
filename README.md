# pptx-copy-ops

`pptx-copy-ops` is a lightweight Python library for copying slides across PowerPoint files using native PPTX XML operations (no HTML conversion).

`pptx-copy-ops` 是一个轻量级 Python 库，使用 PPTX 原生 XML 机制进行跨模板幻灯片复制（不经过 HTML 转换）。

It provides two copy modes:

- `part`: full-slide part-level copy for highest visual fidelity across templates.
- `shape`: blank-slide + shape-level copy fallback.

它提供两种复制模式：

- `part`：整页 part 级复制，优先保证跨模板视觉一致性。
- `shape`：空白页 + shape 级复制，作为兼容回退模式。

## Why / 为什么

Cross-template slide copy can easily break due to:

- dangling relationship IDs
- unregistered slide masters
- `sldMasterId` and `sldLayoutId` collisions

跨模板复制常见的损坏原因包括：

- 关系 ID 悬空（dangling relationship）
- 幻灯片母版未注册
- `sldMasterId` 与 `sldLayoutId` 冲突

This package wraps those low-level OpenXML steps into a stable API.

本项目把这些底层 OpenXML 细节封装成稳定 API，方便直接集成。

## Install / 安装

```bash
pip install pptx-copy-ops
```

## Python API / Python 调用

```python
from pptx_copy_ops import copy_pptx_slides

copy_pptx_slides(
    target_template="/path/to/target-template.pptx",
    sources=[
        ("/path/to/theme1.pptx", 0),   # 0-based
        ("/path/to/theme2.pptx", 12),  # 0-based
    ],
    output_pptx="/path/to/output.pptx",
    mode="part",
)
```

## CLI / 命令行

```bash
python -m pptx_copy_ops.cli \
  --target /path/to/target-template.pptx \
  --source /path/to/theme1.pptx:1 \
  --source /path/to/theme2.pptx:13 \
  --mode part \
  --output /path/to/output.pptx
```

Arguments / 参数说明:

- `--mode`: `part` or `shape`
- `--source`: repeatable, 1-based slide index (`/path/file.pptx:slide_number`)

参数补充：

- `--mode`：可选 `part` 或 `shape`
- `--source`：可重复传入，页码为 1-based（`/path/file.pptx:slide_number`）

## Guarantees / 能力保证

- Part-level copy imports the source slide's part graph, including its slide layout and slide master.
- Part-level copy skips `notesSlide` relation to avoid noisy cross-package note dependencies.
- Imported slide masters are auto-registered in `presentation.xml`.
- Layout IDs are normalized globally and kept disjoint from master IDs.

对应中文：

- `part` 模式会导入源幻灯片的 part 关系图，包括对应版式和幻灯片母版。
- `part` 模式会跳过 `notesSlide` 关系，避免跨包注释页噪声依赖。
- 自动将导入母版注册到 `presentation.xml`。
- 全局规整 `layout id`，并与 `master id` 保持互斥，避免冲突。

## Development / 开发

```bash
cd pptx-copy-ops
pytest -q
python -m build
```

## License / 协议

MIT
