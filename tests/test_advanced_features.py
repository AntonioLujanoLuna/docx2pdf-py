"""Advanced OOXML features and rendering-resource behavior."""

from docx2pdf_py import ConversionOptions
from docx2pdf_py import converter as C
from docx2pdf_py.converter import Converter
from tests.conftest import NS, document


def test_numbering_start_override(make_docx):
    numbering = (
        f'<w:numbering {NS}>'
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
        '<w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        '<w:num w:numId="7"><w:abstractNumId w:val="0"/>'
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/>'
        '</w:lvlOverride></w:num></w:numbering>'
    )
    body = (
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/>'
        '</w:numPr></w:pPr><w:r><w:t>five</w:t></w:r></w:p>'
    )
    with Converter(make_docx(document(body), parts={"word/numbering.xml": numbering})) as conv:
        html = conv.build_html()
    assert "5. <span" in html


def test_tracked_changes_render_insertions_only(make_docx):
    body = (
        '<w:p><w:ins><w:r><w:t>kept</w:t></w:r></w:ins>'
        '<w:del><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>'
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "kept" in html
    assert "removed" not in html


def test_footnotes_and_endnotes(make_docx):
    body = (
        '<w:p><w:r><w:t>body</w:t><w:footnoteReference w:id="2"/>'
        '<w:endnoteReference w:id="3"/></w:r></w:p>'
    )
    footnotes = (
        f'<w:footnotes {NS}><w:footnote w:id="2">'
        '<w:p><w:r><w:t>foot text</w:t></w:r></w:p>'
        '</w:footnote></w:footnotes>'
    )
    endnotes = (
        f'<w:endnotes {NS}><w:endnote w:id="3">'
        '<w:p><w:r><w:t>end text</w:t></w:r></w:p>'
        '</w:endnote></w:endnotes>'
    )
    parts = {"word/footnotes.xml": footnotes, "word/endnotes.xml": endnotes}
    with Converter(make_docx(document(body), parts=parts)) as conv:
        html = conv.build_html()
    assert 'href="#footnote-2"' in html and 'id="footnote-2"' in html
    assert 'href="#endnote-3"' in html and 'id="endnote-3"' in html
    assert "foot text" in html and "end text" in html


def test_textbox_equation_chart_and_object_fallbacks(make_docx):
    math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    chart_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    body = (
        '<w:p><w:r><w:drawing><w:txbxContent><w:p><w:r><w:t>box text</w:t>'
        '</w:r></w:p></w:txbxContent></w:drawing></w:r>'
        f'<m:oMath xmlns:m="{math_ns}"><m:r><m:t>x+1</m:t></m:r></m:oMath>'
        f'<w:r><w:drawing><c:chart xmlns:c="{chart_ns}"/></w:drawing></w:r>'
        '<w:r><w:object/></w:r></w:p>'
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "box text" in html
    assert 'class="equation">x+1' in html
    assert "[Chart]" in html
    assert "[Embedded object]" in html


def test_columns_and_table_pagination_hints(make_docx):
    body = (
        '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
        '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>'
        '<w:tc><w:p><w:r><w:t>header</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>body</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        '<w:sectPr><w:cols w:num="2" w:space="720"/></w:sectPr>'
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "column-count: 2" in html
    assert "column-gap: 36.0pt" in html
    assert "<thead>" in html
    assert 'style="break-inside:avoid"' in html


def test_multiple_sections_get_named_page_geometry(make_docx):
    body = (
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906"/>'
        '<w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720"/>'
        '</w:sectPr></w:pPr><w:r><w:t>landscape section</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>portrait section</w:t></w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "@page section-0 { size: 29.70cm 21.00cm" in html
    assert "@page section-1 { size: 21.00cm 29.70cm" in html
    assert "page: section-0" in html and "page: section-1" in html


def test_rendering_assets_are_written_outside_html(make_docx, tmp_path):
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="image" Target="media/i.png"/>'
        '</Relationships>'
    )
    body = (
        '<w:p><w:r><w:drawing xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        '<wp:inline><a:graphic><a:graphicData><a:blip r:embed="rId1"/>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    )
    parts = {"word/_rels/document.xml.rels": rels, "word/media/i.png": b"png-data"}
    assets = tmp_path / "assets"
    with Converter(make_docx(document(body), parts=parts), asset_directory=assets) as conv:
        html = conv.build_html()
    assert "data:image" not in html
    assert 'src="assets/' in html
    assert len(list(assets.iterdir())) == 1


def test_xml_element_limit(make_docx, monkeypatch):
    monkeypatch.setattr(C, "MAX_XML_ELEMENTS", 3)
    body = '<w:p><w:r><w:t>too many nodes</w:t></w:r></w:p>'
    try:
        Converter(make_docx(document(body)))
    except ValueError as exc:
        assert "XML element count" in str(exc)
    else:
        raise AssertionError("element limit was not enforced")


def test_relationship_target_cannot_escape_package(make_docx):
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="image" Target="../../../outside.png"/>'
        '</Relationships>'
    )
    body = (
        '<w:p><w:r><w:drawing xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        '<wp:inline><a:graphic><a:graphicData><a:blip r:embed="rId1"/>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    )
    with Converter(make_docx(document(body), parts={"word/_rels/document.xml.rels": rels})) as conv:
        try:
            conv.build_html()
        except ValueError as exc:
            assert "escapes OOXML package" in str(exc)
        else:
            raise AssertionError("relationship traversal was accepted")


def test_conversion_options_validate_ranges():
    for kwargs in (
        {"weasyprint_timeout": -1},
        {"native_engine_timeout": -1},
        {"body_line_height": 0},
        {"cell_line_height": float("nan")},
    ):
        try:
            ConversionOptions(**kwargs)
        except ValueError:
            pass
        else:
            raise AssertionError(f"invalid options accepted: {kwargs}")
