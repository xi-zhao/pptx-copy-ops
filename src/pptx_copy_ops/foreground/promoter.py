from __future__ import annotations

import copy
import posixpath
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from .classifier import classify_element
from .dependency_graph import (
    add_relationship,
    allocate_part_name,
    new_relationships_root,
    read_relationships,
    RelationshipInfo,
    relative_target,
    rels_path,
    resolve_target,
)
from .inventory import DML_NS, PML_NS, REL_NS, LayerInventory, q
from .models import ElementClassification, ForegroundCopyPolicy, LayerName


@dataclass
class PromotionReport:
    promoted_elements: list[str]
    copied_parts: list[str]


def _read_entries(path: Path) -> tuple[dict[str, bytes], list[str]]:
    with zipfile.ZipFile(path, "r") as archive:
        names = archive.namelist()
        return {name: archive.read(name) for name in names}, names


def _write_entries(path: Path, entries: dict[str, bytes], ordered_names: list[str]) -> None:
    temp_path = path.with_suffix(path.suffix + ".promote.tmp")
    written: set[str] = set()
    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name in ordered_names:
            if name in entries and name not in written:
                archive.writestr(name, entries[name])
                written.add(name)
        for name, data in entries.items():
            if name not in written:
                archive.writestr(name, data)
                written.add(name)
    temp_path.replace(path)


def _slide_parts_by_order(entries: dict[str, bytes]) -> list[str]:
    presentation = etree.fromstring(entries["ppt/presentation.xml"])
    presentation_rels = etree.fromstring(entries["ppt/_rels/presentation.xml.rels"])
    rid_to_target = {
        rel.get("Id"): rel.get("Target")
        for rel in presentation_rels
        if rel.get("Id") and rel.get("Target")
    }
    slide_parts: list[str] = []
    for slide_node in presentation.findall(f"./{q(PML_NS, 'sldIdLst')}/{q(PML_NS, 'sldId')}"):
        rid = slide_node.get(q(REL_NS, "id"))
        target = rid_to_target.get(rid)
        if target:
            slide_parts.append(resolve_target("ppt/presentation.xml", target))
    return slide_parts


def _source_elements_by_ref(
    source_entries: dict[str, bytes],
    inventory: LayerInventory,
) -> dict[tuple[LayerName, int], etree._Element]:
    needed_parts = {inventory.layout_part, inventory.master_part}
    result: dict[tuple[LayerName, int], etree._Element] = {}
    for layer, part_name in (
        (LayerName.LAYOUT, inventory.layout_part),
        (LayerName.MASTER, inventory.master_part),
    ):
        if part_name not in needed_parts or part_name not in source_entries:
            continue
        root = etree.fromstring(source_entries[part_name])
        sp_tree = root.find(f".//{q(PML_NS, 'spTree')}")
        if sp_tree is None:
            continue
        for index, child in enumerate(list(sp_tree)):
            result[(layer, index)] = child
    return result


def _max_shape_id(slide_root: etree._Element) -> int:
    max_id = 0
    for node in slide_root.findall(f".//{q(PML_NS, 'cNvPr')}"):
        raw_id = node.get("id")
        if raw_id and raw_id.isdigit():
            max_id = max(max_id, int(raw_id))
    return max_id


def _renumber_shape_ids(clone: etree._Element, next_id: int) -> int:
    for node in clone.findall(f".//{q(PML_NS, 'cNvPr')}"):
        node.set("id", str(next_id))
        next_id += 1
    return next_id


def _insert_before_ext_list(sp_tree: etree._Element, clone: etree._Element) -> None:
    for index, child in enumerate(list(sp_tree)):
        if child.tag == q(PML_NS, "extLst"):
            sp_tree.insert(index, clone)
            return
    sp_tree.append(clone)


def _source_relationships(
    source_entries: dict[str, bytes],
    source_part: str,
) -> dict[str, RelationshipInfo]:
    rels_part = rels_path(source_part)
    if rels_part not in source_entries:
        return {}
    _root, rels = read_relationships(source_entries[rels_part])
    return rels


def _copy_internal_part(
    source_entries: dict[str, bytes],
    output_entries: dict[str, bytes],
    existing_names: set[str],
    source_part: str,
    copied_parts: list[str],
) -> str:
    if source_part not in source_entries:
        return source_part
    copied_part = allocate_part_name(source_part, existing_names)
    output_entries[copied_part] = source_entries[source_part]
    copied_parts.append(copied_part)
    return copied_part


def _rewrite_relationships(
    clone: etree._Element,
    *,
    source_part: str,
    source_entries: dict[str, bytes],
    output_entries: dict[str, bytes],
    target_slide_part: str,
    target_rels_root: etree._Element,
    existing_names: set[str],
    copied_parts: list[str],
) -> None:
    source_rels = _source_relationships(source_entries, source_part)
    for element in clone.iter():
        for attr_name in list(element.attrib.keys()):
            if not attr_name.startswith(f"{{{REL_NS}}}"):
                continue
            old_rid = element.attrib.get(attr_name)
            if not old_rid:
                continue
            rel = source_rels.get(old_rid)
            if rel is None:
                del element.attrib[attr_name]
                continue
            if rel.is_external:
                new_rid = add_relationship(
                    target_rels_root,
                    reltype=rel.reltype,
                    target=rel.target,
                    target_mode=rel.target_mode,
                )
                element.set(attr_name, new_rid)
                continue

            source_target_part = resolve_target(source_part, rel.target)
            copied_target_part = _copy_internal_part(
                source_entries,
                output_entries,
                existing_names,
                source_target_part,
                copied_parts,
            )
            new_target = relative_target(target_slide_part, copied_target_part)
            new_rid = add_relationship(
                target_rels_root,
                reltype=rel.reltype,
                target=new_target,
            )
            element.set(attr_name, new_rid)


def promote_foreground_elements(
    *,
    output_pptx: Path | str,
    source_pptx: Path | str,
    inventory: LayerInventory,
    policy: ForegroundCopyPolicy,
    target_slide_index: int = -1,
) -> PromotionReport:
    output_path = Path(output_pptx)
    source_path = Path(source_pptx)
    output_entries, output_order = _read_entries(output_path)
    source_entries, _source_order = _read_entries(source_path)

    slide_parts = _slide_parts_by_order(output_entries)
    if not slide_parts:
        return PromotionReport(promoted_elements=[], copied_parts=[])
    target_slide_part = slide_parts[target_slide_index]
    target_rels_part = rels_path(target_slide_part)
    if target_rels_part in output_entries:
        target_rels_root, _target_rels = read_relationships(output_entries[target_rels_part])
    else:
        target_rels_root = new_relationships_root()

    slide_root = etree.fromstring(output_entries[target_slide_part])
    sp_tree = slide_root.find(f".//{q(PML_NS, 'spTree')}")
    if sp_tree is None:
        return PromotionReport(promoted_elements=[], copied_parts=[])

    source_elements = _source_elements_by_ref(source_entries, inventory)
    existing_names = set(output_entries)
    promoted: list[str] = []
    copied_parts: list[str] = []
    next_shape_id = _max_shape_id(slide_root) + 1

    for element_ref in inventory.elements:
        if element_ref.layer == LayerName.SLIDE:
            continue
        decision = classify_element(element_ref, policy)
        if decision.classification != ElementClassification.FOREGROUND:
            continue
        source_element = source_elements.get((element_ref.layer, element_ref.element_index))
        if source_element is None:
            continue
        clone = copy.deepcopy(source_element)
        next_shape_id = _renumber_shape_ids(clone, next_shape_id)
        _rewrite_relationships(
            clone,
            source_part=element_ref.part_name,
            source_entries=source_entries,
            output_entries=output_entries,
            target_slide_part=target_slide_part,
            target_rels_root=target_rels_root,
            existing_names=existing_names,
            copied_parts=copied_parts,
        )
        _insert_before_ext_list(sp_tree, clone)
        promoted.append(f"{element_ref.layer.value}:{element_ref.name or element_ref.tag}")

    if promoted:
        output_entries[target_slide_part] = etree.tostring(
            slide_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        output_entries[target_rels_part] = etree.tostring(
            target_rels_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        _write_entries(output_path, output_entries, output_order)

    return PromotionReport(promoted_elements=promoted, copied_parts=copied_parts)
