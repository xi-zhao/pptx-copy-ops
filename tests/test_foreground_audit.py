from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from pptx_copy_ops.foreground.audit import audit_pptx_integrity


def test_audit_flags_invalid_chart_external_data(tmp_path: Path) -> None:
    pptx_path = tmp_path / "bad_chart.pptx"
    with zipfile.ZipFile(pptx_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "ppt/charts/chart1.xml",
            """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:externalData r:id="rId3"/>
</c:chartSpace>""",
        )
        archive.writestr(
            "ppt/charts/_rels/chart1.xml.rels",
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
</Relationships>""",
        )

    issues = audit_pptx_integrity(pptx_path, require_openable=False)

    assert any("chart externalData has invalid reltype" in issue for issue in issues)


def test_audit_flags_duplicate_zip_members(tmp_path: Path) -> None:
    pptx_path = tmp_path / "duplicates.pptx"
    with pytest.warns(UserWarning):
        with zipfile.ZipFile(pptx_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("ppt/slides/slide1.xml", "<p:sld xmlns:p='p'/>")
            archive.writestr("ppt/slides/slide1.xml", "<p:sld xmlns:p='p'/>")

    issues = audit_pptx_integrity(pptx_path, require_openable=False)

    assert any("duplicate ZIP members" in issue for issue in issues)


def test_audit_flags_missing_relationship_target(tmp_path: Path) -> None:
    pptx_path = tmp_path / "missing_rel_target.pptx"
    with zipfile.ZipFile(pptx_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "ppt/slides/slide1.xml",
            """<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree/></p:cSld>
</p:sld>""",
        )
        archive.writestr(
            "ppt/slides/_rels/slide1.xml.rels",
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/missing.png"/>
</Relationships>""",
        )

    issues = audit_pptx_integrity(pptx_path, require_openable=False)

    assert any("missing relationship target" in issue for issue in issues)


def test_audit_flags_relationship_target_without_content_type(tmp_path: Path) -> None:
    pptx_path = tmp_path / "missing_content_type.pptx"
    with zipfile.ZipFile(pptx_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>""",
        )
        archive.writestr(
            "ppt/slides/slide1.xml",
            """<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree/></p:cSld>
</p:sld>""",
        )
        archive.writestr(
            "ppt/slides/_rels/slide1.xml.rels",
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>""",
        )
        archive.writestr("ppt/charts/chart1.xml", "<c:chartSpace xmlns:c='c'/>")

    issues = audit_pptx_integrity(pptx_path, require_openable=False)

    assert any("missing content type for relationship target" in issue for issue in issues)
