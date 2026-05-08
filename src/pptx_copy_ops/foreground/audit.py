from __future__ import annotations

import posixpath
import re
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree
from pptx import Presentation

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_REL_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def rels_source_part(rels_name: str) -> str | None:
    if rels_name == "_rels/.rels":
        return ""
    if "/_rels/" not in rels_name or not rels_name.endswith(".rels"):
        return None
    base_dir, filename = rels_name.rsplit("/_rels/", 1)
    return f"{base_dir}/{filename[:-5]}"


def rels_part_for(source_part: str) -> str:
    if not source_part:
        return "_rels/.rels"
    return posixpath.join(
        posixpath.dirname(source_part),
        "_rels",
        f"{posixpath.basename(source_part)}.rels",
    )


def resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    base_dir = posixpath.dirname(source_part) if source_part else ""
    return posixpath.normpath(posixpath.join(base_dir, target))


def read_relationships(
    archive: zipfile.ZipFile,
    names: set[str],
) -> dict[str, dict[str, tuple[str, str, str]]]:
    relationships: dict[str, dict[str, tuple[str, str, str]]] = {}
    for name in names:
        if not name.endswith(".rels"):
            continue
        source_part = rels_source_part(name)
        if source_part is None:
            continue
        root = etree.fromstring(archive.read(name))
        rels: dict[str, tuple[str, str, str]] = {}
        for rel in root:
            rid = rel.get("Id")
            if not rid:
                continue
            rels[rid] = (
                rel.get("Type", ""),
                rel.get("Target", ""),
                rel.get("TargetMode", ""),
            )
        relationships[source_part] = rels
    return relationships


def _content_type_maps(root: etree._Element | None) -> tuple[set[str], set[str]]:
    if root is None:
        return set(), set()
    overrides = {
        (override.get("PartName") or "").lstrip("/")
        for override in root.findall(f"{{{CONTENT_TYPES_NS}}}Override")
        if override.get("PartName") and override.get("ContentType")
    }
    defaults = {
        (default.get("Extension") or "").lower()
        for default in root.findall(f"{{{CONTENT_TYPES_NS}}}Default")
        if default.get("Extension") and default.get("ContentType")
    }
    return overrides, defaults


def _has_content_type(part_name: str, overrides: set[str], defaults: set[str]) -> bool:
    if part_name in overrides:
        return True
    extension = Path(part_name).suffix.lstrip(".").lower()
    return bool(extension and extension in defaults)


def audit_pptx_integrity(path: Path | str, require_openable: bool = True) -> list[str]:
    pptx_path = Path(path)
    issues: list[str] = []

    if require_openable:
        try:
            prs = Presentation(str(pptx_path))
            if len(prs.slides) == 0:
                issues.append("presentation has no slides")
        except Exception as exc:
            issues.append(f"python-pptx cannot open presentation: {type(exc).__name__}: {exc}")

    with zipfile.ZipFile(pptx_path) as archive:
        names = archive.namelist()
        name_set = set(names)

        duplicate_members = [name for name, count in Counter(names).items() if count > 1]
        if duplicate_members:
            issues.append(f"duplicate ZIP members: {duplicate_members[:20]}")

        xml_roots: dict[str, etree._Element] = {}
        for name in names:
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                xml_roots[name] = etree.fromstring(archive.read(name))
            except Exception as exc:
                issues.append(f"invalid XML {name}: {type(exc).__name__}: {exc}")

        relationships = read_relationships(archive, name_set)
        content_types_root = xml_roots.get("[Content_Types].xml")
        if content_types_root is None:
            issues.append("missing [Content_Types].xml")
        content_type_overrides, content_type_defaults = _content_type_maps(content_types_root)

        for rels_name, root in xml_roots.items():
            if not rels_name.endswith(".rels"):
                continue
            source_part = rels_source_part(rels_name)
            if source_part is None:
                continue
            rel_ids: list[str] = []
            for rel in root:
                rid = rel.get("Id")
                if rid:
                    rel_ids.append(rid)
                target = rel.get("Target")
                if not target or rel.get("TargetMode") == "External":
                    continue
                resolved = resolve_target(source_part, target)
                if resolved not in name_set:
                    issues.append(
                        f"missing relationship target: {rels_name} {rid} -> {target} ({resolved})"
                    )
                elif not _has_content_type(
                    resolved,
                    content_type_overrides,
                    content_type_defaults,
                ):
                    issues.append(
                        f"missing content type for relationship target: {rels_name} {rid} -> {resolved}"
                    )
            duplicate_rel_ids = [rid for rid, count in Counter(rel_ids).items() if count > 1]
            if duplicate_rel_ids:
                issues.append(f"duplicate relationship ids in {rels_name}: {duplicate_rel_ids}")

        for part_name, root in xml_roots.items():
            if part_name.endswith(".rels") or part_name == "[Content_Types].xml":
                continue
            part_rel_ids = set(relationships.get(part_name, {}))
            for element in root.iter():
                for attr_name, value in element.attrib.items():
                    if not value:
                        continue
                    is_rel_attr = attr_name.startswith(f"{{{REL_NS}}}")
                    is_rid_like = re.fullmatch(r"rId\d+", value) is not None
                    if (is_rel_attr or is_rid_like) and value not in part_rel_ids:
                        issues.append(f"dangling relationship id: {part_name} {attr_name}={value}")

        valid_external_data_reltypes = {
            f"{OFFICE_REL_PREFIX}package",
            f"{OFFICE_REL_PREFIX}oleObject",
        }
        for part_name, root in xml_roots.items():
            if not (part_name.startswith("ppt/charts/chart") and part_name.endswith(".xml")):
                continue
            rel_map = relationships.get(part_name, {})
            for node in root.findall(f".//{q(CHART_NS, 'externalData')}"):
                rid = node.get(q(REL_NS, "id"))
                rel_type, target, target_mode = rel_map.get(rid or "", ("", "", ""))
                if not rid:
                    issues.append(f"chart externalData missing r:id: {part_name}")
                elif not rel_type:
                    issues.append(f"chart externalData references missing rel: {part_name} {rid}")
                elif target_mode == "External":
                    issues.append(f"chart externalData uses external relationship: {part_name} {rid} -> {target}")
                elif rel_type not in valid_external_data_reltypes:
                    issues.append(f"chart externalData has invalid reltype: {part_name} {rid} {rel_type}")

            chart_rels_name = rels_part_for(part_name)
            for rid, (rel_type, target, target_mode) in rel_map.items():
                if target_mode == "External" and rel_type == f"{OFFICE_REL_PREFIX}oleObject":
                    issues.append(f"chart external OLE relationship: {chart_rels_name} {rid} -> {target}")

    return issues
