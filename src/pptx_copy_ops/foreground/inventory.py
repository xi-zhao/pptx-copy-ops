from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from .models import ElementRef, LayerName

PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SLIDE_LAYOUT_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
SLIDE_MASTER_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


@dataclass(frozen=True)
class LayerInventory:
    slide_part: str
    layout_part: str
    master_part: str
    elements: list[ElementRef]


def _rels_path(part_name: str) -> str:
    return posixpath.join(
        posixpath.dirname(part_name),
        "_rels",
        f"{posixpath.basename(part_name)}.rels",
    )


def _resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _shape_text(node: etree._Element) -> str:
    return "".join("".join(text.itertext()) for text in node.findall(f".//{q(DML_NS, 't')}")).strip()


def _shape_name(node: etree._Element) -> str:
    c_nv_pr = node.find(f".//{q(PML_NS, 'cNvPr')}")
    return c_nv_pr.get("name", "") if c_nv_pr is not None else ""


def _shape_extents(node: etree._Element) -> tuple[int, int, int, int] | None:
    xfrm = node.find(f".//{q(DML_NS, 'xfrm')}")
    if xfrm is None:
        return None
    off = xfrm.find(q(DML_NS, "off"))
    ext = xfrm.find(q(DML_NS, "ext"))
    if off is None or ext is None:
        return None
    return (
        int(off.get("x", "0")),
        int(off.get("y", "0")),
        int(ext.get("cx", "0")),
        int(ext.get("cy", "0")),
    )


def _collect_elements(archive: zipfile.ZipFile, part_name: str, layer: LayerName) -> list[ElementRef]:
    root = etree.fromstring(archive.read(part_name))
    sp_tree = root.find(f".//{q(PML_NS, 'spTree')}")
    if sp_tree is None:
        return []

    elements: list[ElementRef] = []
    for index, child in enumerate(list(sp_tree)):
        if child.tag.endswith("}nvGrpSpPr") or child.tag.endswith("}grpSpPr"):
            continue
        elements.append(
            ElementRef(
                layer=layer,
                part_name=part_name,
                element_index=index,
                tag=etree.QName(child).localname,
                name=_shape_name(child),
                text=_shape_text(child),
                extents=_shape_extents(child),
            )
        )
    return elements


def _relationship_target(
    archive: zipfile.ZipFile,
    rels_part: str,
    reltype: str,
) -> str:
    root = etree.fromstring(archive.read(rels_part))
    for rel in root:
        if rel.get("Type") == reltype and rel.get("Target"):
            return rel.get("Target") or ""
    raise ValueError(f"missing relationship {reltype} in {rels_part}")


def collect_layer_inventory(source_pptx: Path | str, slide_index: int) -> LayerInventory:
    path = Path(source_pptx).expanduser().resolve()
    with zipfile.ZipFile(path) as archive:
        presentation = etree.fromstring(archive.read("ppt/presentation.xml"))
        presentation_rels = etree.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
        rid_to_target = {
            rel.get("Id"): rel.get("Target")
            for rel in presentation_rels
            if rel.get("Id") and rel.get("Target")
        }
        slide_nodes = presentation.findall(f"./{q(PML_NS, 'sldIdLst')}/{q(PML_NS, 'sldId')}")
        if slide_index < 0 or slide_index >= len(slide_nodes):
            raise IndexError(f"slide_index {slide_index} out of range")

        slide_rid = slide_nodes[slide_index].get(q(REL_NS, "id"))
        slide_target = rid_to_target.get(slide_rid)
        if not slide_target:
            raise ValueError(f"missing slide relationship target for {slide_rid}")

        slide_part = _resolve_target("ppt/presentation.xml", slide_target)
        layout_target = _relationship_target(archive, _rels_path(slide_part), SLIDE_LAYOUT_RELTYPE)
        layout_part = _resolve_target(slide_part, layout_target)
        master_target = _relationship_target(archive, _rels_path(layout_part), SLIDE_MASTER_RELTYPE)
        master_part = _resolve_target(layout_part, master_target)

        elements: list[ElementRef] = []
        elements.extend(_collect_elements(archive, slide_part, LayerName.SLIDE))
        elements.extend(_collect_elements(archive, layout_part, LayerName.LAYOUT))
        elements.extend(_collect_elements(archive, master_part, LayerName.MASTER))

    return LayerInventory(
        slide_part=slide_part,
        layout_part=layout_part,
        master_part=master_part,
        elements=elements,
    )
