from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
OFFICE_REL_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


@dataclass
class SanitizationReport:
    sanitized_parts: list[str] = field(default_factory=list)


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _chart_part_for_rels(rels_name: str) -> str:
    return posixpath.normpath(
        posixpath.join(
            "ppt/charts",
            posixpath.basename(rels_name).removesuffix(".rels"),
        )
    )


def sanitize_pptx_package(pptx_path: Path | str) -> SanitizationReport:
    source_path = Path(pptx_path)
    temp_path = source_path.with_suffix(source_path.suffix + ".sanitize.tmp")
    report = SanitizationReport()
    ext_lst_xpath = f".//{q(CHART_NS, 'extLst')}"
    external_data_xpath = f".//{q(CHART_NS, 'externalData')}"
    valid_external_data_reltypes = {
        f"{OFFICE_REL_PREFIX}package",
        f"{OFFICE_REL_PREFIX}oleObject",
    }

    chart_rel_types: dict[str, dict[str, tuple[str, str]]] = {}
    with zipfile.ZipFile(source_path, "r") as archive:
        for name in archive.namelist():
            if not (name.startswith("ppt/charts/_rels/chart") and name.endswith(".xml.rels")):
                continue
            try:
                root = etree.fromstring(archive.read(name))
            except Exception:
                continue
            chart_rel_types[_chart_part_for_rels(name)] = {
                rel.attrib.get("Id", ""): (
                    rel.attrib.get("Type", ""),
                    rel.attrib.get("TargetMode", ""),
                )
                for rel in root
                if rel.attrib.get("Id")
            }

    changed = False
    with zipfile.ZipFile(source_path, "r") as zin, zipfile.ZipFile(
        temp_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
    ) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("ppt/charts/chart") and info.filename.endswith(".xml"):
                try:
                    root = etree.fromstring(data)
                except Exception:
                    zout.writestr(info, data)
                    continue

                removed = 0
                for node in root.findall(ext_lst_xpath):
                    parent = node.getparent()
                    if parent is not None:
                        parent.remove(node)
                        removed += 1

                rel_map = chart_rel_types.get(info.filename, {})
                for node in root.findall(external_data_xpath):
                    rid = node.attrib.get(q(REL_NS, "id"))
                    rel_type, target_mode = rel_map.get(rid or "", ("", ""))
                    if (
                        not rid
                        or not rel_type
                        or target_mode == "External"
                        or rel_type not in valid_external_data_reltypes
                    ):
                        parent = node.getparent()
                        if parent is not None:
                            parent.remove(node)
                            removed += 1

                if removed:
                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                    changed = True
                    report.sanitized_parts.append(info.filename)

            elif info.filename.startswith("ppt/charts/_rels/chart") and info.filename.endswith(".xml.rels"):
                try:
                    root = etree.fromstring(data)
                except Exception:
                    zout.writestr(info, data)
                    continue

                removed = 0
                for rel in list(root):
                    rel_type = rel.attrib.get("Type", "")
                    target_mode = rel.attrib.get("TargetMode", "")
                    if target_mode == "External" and rel_type == f"{OFFICE_REL_PREFIX}oleObject":
                        root.remove(rel)
                        removed += 1

                if removed:
                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                    changed = True
                    report.sanitized_parts.append(info.filename)

            zout.writestr(info, data)

    if changed:
        temp_path.replace(source_path)
    else:
        temp_path.unlink(missing_ok=True)
    return report
