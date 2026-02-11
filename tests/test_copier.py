from __future__ import annotations

import base64
import os
import posixpath
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pptx_copy_ops.copier import SlideCopier, SlideSpec  # noqa: E402

_ONE_PIXEL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jM5cAAAAASUVORK5CYII="
)


def _build_target_template(path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    if slide.shapes.title is not None:
        slide.shapes.title.text = "Target Template"
    prs.save(path)


def _build_source_deck(path: Path, title: str, body: str) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    if slide.shapes.title is not None:
        slide.shapes.title.text = title
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = body

    image_path = path.with_suffix(".png")
    image_path.write_bytes(_ONE_PIXEL_PNG)
    slide.shapes.add_picture(str(image_path), Inches(9.0), Inches(0.2), Inches(0.35), Inches(0.35))
    prs.save(path)
    image_path.unlink(missing_ok=True)


def _slide_texts(slide) -> list[str]:
    texts: list[str] = []
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if text:
            texts.append(text)
    return texts


def _assert_no_dangling_slide_relationship_refs(pptx_path: Path) -> None:
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rel_prefix = f"{{{rel_ns}}}"
    with zipfile.ZipFile(pptx_path) as archive:
        names = set(archive.namelist())
        slide_parts = sorted(
            name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        )
        for slide_part in slide_parts:
            rel_part = posixpath.join(
                posixpath.dirname(slide_part),
                "_rels",
                f"{posixpath.basename(slide_part)}.rels",
            )
            rel_ids: set[str] = set()
            if rel_part in names:
                rel_root = ET.fromstring(archive.read(rel_part))
                for rel in rel_root:
                    rel_id = rel.attrib.get("Id")
                    if rel_id:
                        rel_ids.add(rel_id)
            slide_root = ET.fromstring(archive.read(slide_part))
            for element in slide_root.iter():
                for attr_name, attr_value in element.attrib.items():
                    if not attr_name.startswith(rel_prefix):
                        continue
                    assert attr_value in rel_ids, (
                        f"dangling relationship reference {attr_name}={attr_value} in {slide_part}, "
                        f"missing from {rel_part}"
                    )


def _assert_master_ids_disjoint_from_layout_ids(pptx_path: Path) -> None:
    p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    with zipfile.ZipFile(pptx_path) as archive:
        presentation_root = ET.fromstring(archive.read("ppt/presentation.xml"))
        master_ids = {
            int(node.attrib["id"])
            for node in presentation_root.findall(f"./{{{p_ns}}}sldMasterIdLst/{{{p_ns}}}sldMasterId")
            if node.attrib.get("id")
        }

        layout_ids: set[int] = set()
        master_parts = sorted(
            name
            for name in archive.namelist()
            if name.startswith("ppt/slideMasters/slideMaster") and name.endswith(".xml")
        )
        for master_part in master_parts:
            master_root = ET.fromstring(archive.read(master_part))
            for node in master_root.findall(f".//{{{p_ns}}}sldLayoutId"):
                raw_id = node.attrib.get("id")
                if raw_id:
                    layout_ids.add(int(raw_id))

        overlap = sorted(master_ids & layout_ids)
        assert not overlap, f"sldMasterId/@id conflicts with sldLayoutId/@id: {overlap}"


def test_part_mode_copy_two_slides_and_id_invariants(tmp_path: Path) -> None:
    target = tmp_path / "target.pptx"
    src1 = tmp_path / "source1.pptx"
    src2 = tmp_path / "source2.pptx"
    output = tmp_path / "out_part.pptx"

    _build_target_template(target)
    _build_source_deck(src1, title="Source A", body="Body A")
    _build_source_deck(src2, title="Source B", body="Body B")

    copier = SlideCopier(target)
    copier.copy_slides(
        [
            SlideSpec(src1, 0),
            SlideSpec(src2, 0),
        ],
        mode="part",
    )
    copier.save(output)

    generated = Presentation(output)
    assert len(generated.slides) == 2
    assert "Source A" in _slide_texts(generated.slides[0])
    assert "Source B" in _slide_texts(generated.slides[1])

    _assert_no_dangling_slide_relationship_refs(output)
    _assert_master_ids_disjoint_from_layout_ids(output)


def test_shape_mode_copy_two_slides_text_matches(tmp_path: Path) -> None:
    target = tmp_path / "target.pptx"
    src1 = tmp_path / "source1.pptx"
    src2 = tmp_path / "source2.pptx"
    output = tmp_path / "out_shape.pptx"

    _build_target_template(target)
    _build_source_deck(src1, title="Shape Source A", body="Shape Body A")
    _build_source_deck(src2, title="Shape Source B", body="Shape Body B")

    copier = SlideCopier(target)
    copier.copy_slides(
        [
            SlideSpec(src1, 0),
            SlideSpec(src2, 0),
        ],
        mode="shape",
    )
    copier.save(output)

    generated = Presentation(output)
    assert len(generated.slides) == 2
    assert "Shape Source A" in _slide_texts(generated.slides[0])
    assert "Shape Source B" in _slide_texts(generated.slides[1])

    _assert_no_dangling_slide_relationship_refs(output)


def test_cli_smoke(tmp_path: Path) -> None:
    target = tmp_path / "target.pptx"
    src1 = tmp_path / "source1.pptx"
    src2 = tmp_path / "source2.pptx"
    output = tmp_path / "out_cli.pptx"

    _build_target_template(target)
    _build_source_deck(src1, title="CLI A", body="CLI Body A")
    _build_source_deck(src2, title="CLI B", body="CLI Body B")

    env = os.environ.copy()
    env["PYTHONPATH"] = str(SRC) + os.pathsep + env.get("PYTHONPATH", "")
    cmd = [
        sys.executable,
        "-m",
        "pptx_copy_ops.cli",
        "--target",
        str(target),
        "--source",
        f"{src1}:1",
        "--source",
        f"{src2}:1",
        "--mode",
        "part",
        "--output",
        str(output),
    ]
    result = subprocess.run(cmd, text=True, capture_output=True, env=env)
    assert result.returncode == 0, f"stdout={result.stdout}\nstderr={result.stderr}"
    assert output.exists()
