from __future__ import annotations

import posixpath
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation

from pptx_copy_ops.foreground import CopyPolicy, ForegroundCopyPolicy, ForegroundCopyRequest
from pptx_copy_ops.foreground.audit import audit_pptx_integrity
from pptx_copy_ops.foreground.engine import copy_foreground_slide
from pptx_copy_ops.foreground.inventory import PML_NS, REL_NS, collect_layer_inventory, q


def _slide_texts(path: Path) -> list[str]:
    prs = Presentation(str(path))
    texts: list[str] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text = shape.text_frame.text.strip()
                if text:
                    texts.append(text)
    return texts


def _rels_path(part_name: str) -> str:
    return posixpath.join(
        posixpath.dirname(part_name),
        "_rels",
        f"{posixpath.basename(part_name)}.rels",
    )


def _inject_master_shapes(path: Path) -> None:
    inventory = collect_layer_inventory(path, 0)
    temp_path = path.with_suffix(".inject.tmp")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(
        temp_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
    ) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == inventory.master_part:
                root = etree.fromstring(data)
                sp_tree = root.find(f".//{q(PML_NS, 'spTree')}")
                assert sp_tree is not None
                sp_tree.append(
                    etree.fromstring(
                        f"""
<p:sp xmlns:p="{PML_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="9001" name="MASTER_BACKGROUND_MARKER"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
  </p:spPr>
</p:sp>
""".strip()
                    )
                )
                sp_tree.append(
                    etree.fromstring(
                        f"""
<p:sp xmlns:p="{PML_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="9002" name="MASTER_FOREGROUND_MARKER"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="400000" y="200000"/><a:ext cx="2200000" cy="350000"/></a:xfrm>
    <a:solidFill><a:srgbClr val="111111"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:r><a:t>MASTER_FOREGROUND_MARKER</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
""".strip()
                    )
                )
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(info, data)
    temp_path.replace(path)


def test_foreground_promote_keeps_master_foreground_and_slide_text(tmp_path: Path) -> None:
    source = tmp_path / "source.pptx"
    target = tmp_path / "target.pptx"
    output = tmp_path / "output.pptx"

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(500000, 700000, 3000000, 400000).text = "SLIDE_FOREGROUND_MARKER"
    prs.save(source)
    _inject_master_shapes(source)
    Presentation().save(target)

    result = copy_foreground_slide(
        ForegroundCopyRequest(
            source_pptx=source,
            slide_index=0,
            target_pptx=target,
            policy=ForegroundCopyPolicy(copy_policy=CopyPolicy.FOREGROUND_PROMOTE),
        ),
        output_pptx=output,
    )

    texts = _slide_texts(output)
    assert "SLIDE_FOREGROUND_MARKER" in texts
    assert "MASTER_FOREGROUND_MARKER" in texts
    assert "MASTER_BACKGROUND_MARKER" not in "\n".join(texts)
    assert result.trace.promoted_elements
    assert result.trace.audit_issues == []
    assert audit_pptx_integrity(output) == []


def test_exact_part_copy_keeps_source_master_registration(tmp_path: Path) -> None:
    source = tmp_path / "source_exact.pptx"
    target = tmp_path / "target_exact.pptx"
    output = tmp_path / "output_exact.pptx"

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(500000, 700000, 3000000, 400000).text = "EXACT_COPY_MARKER"
    prs.save(source)
    Presentation().save(target)

    result = copy_foreground_slide(
        ForegroundCopyRequest(
            source_pptx=source,
            slide_index=0,
            target_pptx=target,
            policy=ForegroundCopyPolicy(copy_policy=CopyPolicy.EXACT_PART_COPY),
        ),
        output_pptx=output,
    )

    assert "EXACT_COPY_MARKER" in _slide_texts(output)
    assert result.trace.audit_issues == []

    with zipfile.ZipFile(output) as archive:
        presentation = etree.fromstring(archive.read("ppt/presentation.xml"))
        master_entries = presentation.findall(f"./{q(PML_NS, 'sldMasterIdLst')}/{q(PML_NS, 'sldMasterId')}")
        slide_nodes = presentation.findall(f"./{q(PML_NS, 'sldIdLst')}/{q(PML_NS, 'sldId')}")
        assert master_entries
        assert slide_nodes[-1].get(q(REL_NS, "id"))

    assert audit_pptx_integrity(output) == []
