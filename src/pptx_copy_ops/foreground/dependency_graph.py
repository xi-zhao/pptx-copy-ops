from __future__ import annotations

import posixpath
from dataclasses import dataclass
from pathlib import PurePosixPath

from lxml import etree

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


@dataclass(frozen=True)
class RelationshipInfo:
    rid: str
    reltype: str
    target: str
    target_mode: str = ""

    @property
    def is_external(self) -> bool:
        return self.target_mode == "External"


def rels_path(part_name: str) -> str:
    return posixpath.join(
        posixpath.dirname(part_name),
        "_rels",
        f"{posixpath.basename(part_name)}.rels",
    )


def resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relative_target(from_part: str, to_part: str) -> str:
    return posixpath.relpath(to_part, posixpath.dirname(from_part))


def read_relationships(data: bytes) -> tuple[etree._Element, dict[str, RelationshipInfo]]:
    root = etree.fromstring(data)
    rels: dict[str, RelationshipInfo] = {}
    for rel in root:
        rid = rel.get("Id")
        if not rid:
            continue
        rels[rid] = RelationshipInfo(
            rid=rid,
            reltype=rel.get("Type", ""),
            target=rel.get("Target", ""),
            target_mode=rel.get("TargetMode", ""),
        )
    return root, rels


def new_relationships_root() -> etree._Element:
    return etree.Element(f"{{{PKG_REL_NS}}}Relationships", nsmap={None: PKG_REL_NS})


def next_rid(root: etree._Element) -> str:
    used: set[int] = set()
    for rel in root:
        rid = rel.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            used.add(int(rid[3:]))
    candidate = 1
    while candidate in used:
        candidate += 1
    return f"rId{candidate}"


def add_relationship(
    root: etree._Element,
    *,
    reltype: str,
    target: str,
    target_mode: str = "",
) -> str:
    rid = next_rid(root)
    rel = etree.SubElement(root, f"{{{PKG_REL_NS}}}Relationship")
    rel.set("Id", rid)
    rel.set("Type", reltype)
    rel.set("Target", target)
    if target_mode:
        rel.set("TargetMode", target_mode)
    return rid


def allocate_part_name(source_part: str, existing_names: set[str]) -> str:
    part = PurePosixPath(source_part)
    stem = part.stem
    suffix = part.suffix
    parent = str(part.parent)
    candidate = source_part
    if candidate not in existing_names:
        existing_names.add(candidate)
        return candidate

    for index in range(1, 100000):
        candidate = posixpath.join(parent, f"{stem}_{index}{suffix}")
        if candidate not in existing_names:
            existing_names.add(candidate)
            return candidate
    raise RuntimeError(f"cannot allocate PPTX part name for {source_part}")
