from __future__ import annotations

import zipfile
from pathlib import Path

from pptx_copy_ops.foreground.sanitizer import sanitize_pptx_package


def test_sanitizer_removes_invalid_chart_external_data(tmp_path: Path) -> None:
    pptx_path = tmp_path / "invalid_chart.pptx"
    with zipfile.ZipFile(pptx_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "ppt/charts/chart1.xml",
            """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:externalData r:id="rId3"/>
  <c:extLst><c:ext uri="{test}"/></c:extLst>
</c:chartSpace>""",
        )
        archive.writestr(
            "ppt/charts/_rels/chart1.xml.rels",
            """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
</Relationships>""",
        )

    report = sanitize_pptx_package(pptx_path)

    with zipfile.ZipFile(pptx_path) as archive:
        chart_xml = archive.read("ppt/charts/chart1.xml").decode("utf-8")
        rels_xml = archive.read("ppt/charts/_rels/chart1.xml.rels").decode("utf-8")
    assert "externalData" not in chart_xml
    assert "extLst" not in chart_xml
    assert "chartStyle" in rels_xml
    assert "ppt/charts/chart1.xml" in report.sanitized_parts


def test_sanitizer_removes_external_chart_ole_relationship(tmp_path: Path) -> None:
    pptx_path = tmp_path / "external_ole.pptx"
    with zipfile.ZipFile(pptx_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "ppt/charts/chart1.xml",
            """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:externalData r:id="rId1"/>
</c:chartSpace>""",
        )
        archive.writestr(
            "ppt/charts/_rels/chart1.xml.rels",
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="file:///C:/Users/source.xlsx" TargetMode="External"/>
</Relationships>""",
        )

    report = sanitize_pptx_package(pptx_path)

    with zipfile.ZipFile(pptx_path) as archive:
        chart_xml = archive.read("ppt/charts/chart1.xml").decode("utf-8")
        rels_xml = archive.read("ppt/charts/_rels/chart1.xml.rels").decode("utf-8")
    assert "externalData" not in chart_xml
    assert "oleObject" not in rels_xml
    assert "ppt/charts/chart1.xml" in report.sanitized_parts
    assert "ppt/charts/_rels/chart1.xml.rels" in report.sanitized_parts
