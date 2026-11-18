from pathlib import Path
import csv, io, json, re, zipfile
from typing import Any
from xml.sax.saxutils import escape
import pptx_writer


def safe_title(title, fallback='Document'):
    return ''.join(ch if ch.isalnum() or ch in ' ._-' else '_' for ch in title).strip(' ._') or fallback

ODF_OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
ODF_TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
ODF_TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
ODF_DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
ODF_PRESENTATION_NS = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
ODF_STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
ODF_FO_NS = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
ODF_MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
ODF_VERSION = "1.2"
ODF_MIMETYPES = {
    "odt": "application/vnd.oasis.opendocument.text",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "odp": "application/vnd.oasis.opendocument.presentation",
}


def normalize_extension(value):
    ext = str(value).lower().strip().lstrip('.')
    if ext not in {'odt','ods','odp','docx','xlsx','pptx'}:
        raise ValueError('Unsupported Office format')
    return ext

def template_bytes(kind: str, ext: str, title: str, content: str) -> bytes:
    ext = normalize_extension(ext or "md")
    if not title and not content and ext in ODF_MIMETYPES:
        body = {
            'odt': '<office:text><text:p/></office:text>',
            'ods': '<office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet>',
            'odp': '<office:presentation><draw:page draw:name="Slide 1" draw:master-page-name="Default"/></office:presentation>',
        }[ext]
        return _odf_package(ext, f'<office:document-content {_odf_content_namespaces()} office:version="{ODF_VERSION}"><office:body>{body}</office:body></office:document-content>')
    if ext == "md":
        return _markdown(title, content).encode("utf-8")
    if ext == "txt":
        return (str(content or "") or str(title or "")).encode("utf-8")
    if ext == "odt":
        return odt_bytes(title, content)
    if ext == "ods":
        return ods_bytes(title, content)
    if ext == "odp":
        return odp_bytes(title, content)
    if ext == "docx":
        return _docx(title, content)
    if ext == "xlsx":
        return _xlsx(title, content)
    if ext == "pptx":
        return _pptx(title, content)
    raise ValueError(ext)


def _markdown(title: str, content: str) -> str:
    text = str(content or "").strip()
    if text:
        return text if text.startswith("#") else f"# {title}\n\n{text}\n"
    return f"# {title}\n"


def _zip_bytes(files_map: dict[str, str | bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, value in files_map.items():
            data = value.encode("utf-8") if isinstance(value, str) else value
            archive.writestr(name, data)
    return buffer.getvalue()


def odf_zip_bytes(ext: str, files_map: dict[str, str | bytes]) -> bytes:
    ext = normalize_extension(ext)
    if ext not in ODF_MIMETYPES:
        raise ValueError(f"Unsupported ODF format: {ext}")
    media_type = ODF_MIMETYPES[ext]
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("mimetype", media_type, compress_type=zipfile.ZIP_STORED)
        for name, value in files_map.items():
            if name == "mimetype":
                continue
            data = value.encode("utf-8") if isinstance(value, str) else value
            archive.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
    return buffer.getvalue()


def odt_bytes(title: str, content: str) -> bytes:
    return odt_bytes_from_paragraphs(_document_lines(title, content))


def odt_bytes_from_paragraphs(paragraphs: list[str]) -> bytes:
    lines = [str(line) for line in paragraphs] or [""]
    body = "\n".join(_odt_paragraph(line, index == 0) for index, line in enumerate(lines))
    return _odf_package(
        "odt",
        f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content {_odf_content_namespaces()} office:version="{ODF_VERSION}">
  <office:body>
    <office:text>
      {body}
    </office:text>
  </office:body>
</office:document-content>
""",
    )


def ods_bytes(title: str, content: str) -> bytes:
    return ods_bytes_from_sheets([{"name": "Sheet1", "rows": _xlsx_rows(title, content)}])


def ods_bytes_from_sheets(sheets: list[dict[str, Any]]) -> bytes:
    normalized = []
    for index, sheet in enumerate(sheets or []):
        name = safe_title(str(sheet.get("name") or f"Sheet{index + 1}"), f"Sheet{index + 1}")[:31] or f"Sheet{index + 1}"
        rows = sheet.get("rows") or []
        normalized.append({"name": name, "rows": rows})
    if not normalized:
        normalized = [{"name": "Sheet1", "rows": [["Spreadsheet"]]}]

    tables = "\n".join(
        f"""<table:table table:name="{escape(sheet['name'])}">
          {''.join(_ods_row(row) for row in sheet['rows'])}
        </table:table>"""
        for sheet in normalized
    )
    return _odf_package(
        "ods",
        f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content {_odf_content_namespaces()} office:version="{ODF_VERSION}">
  <office:body>
    <office:spreadsheet>
      {tables}
    </office:spreadsheet>
  </office:body>
</office:document-content>
""",
    )


def odp_bytes(title: str, content: str) -> bytes:
    return odp_bytes_from_slides(pptx_writer.slides_from_text(title, content))


def odp_bytes_from_slides(slides: list[dict[str, Any]]) -> bytes:
    normalized = pptx_writer.normalize_slides(slides)
    if not normalized:
        normalized = [{"title": "Presentation", "bullets": []}]
    pages = "\n".join(_odp_page(slide, index) for index, slide in enumerate(normalized, start=1))
    return _odf_package(
        "odp",
        f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content {_odf_content_namespaces()} office:version="{ODF_VERSION}">
  <office:body>
    <office:presentation>
      {pages}
    </office:presentation>
  </office:body>
</office:document-content>
""",
    )


def _document_lines(title: str, content: str) -> list[str]:
    lines = [str(title or "Document").strip() or "Document"]
    lines.extend(line.rstrip() for line in str(content or "").splitlines() if line.strip())
    if len(lines) == 1:
        lines.append("")
    return lines


def _odf_package(ext: str, content_xml: str) -> bytes:
    return odf_zip_bytes(
        ext,
        {
            "content.xml": content_xml,
            "styles.xml": _odf_styles_xml(),
            "meta.xml": _odf_meta_xml(),
            "settings.xml": _odf_settings_xml(),
            "META-INF/manifest.xml": _odf_manifest_xml(ODF_MIMETYPES[ext]),
        },
    )


def _odf_content_namespaces() -> str:
    return (
        f'xmlns:office="{ODF_OFFICE_NS}" '
        f'xmlns:text="{ODF_TEXT_NS}" '
        f'xmlns:table="{ODF_TABLE_NS}" '
        f'xmlns:draw="{ODF_DRAW_NS}" '
        f'xmlns:presentation="{ODF_PRESENTATION_NS}" '
        f'xmlns:style="{ODF_STYLE_NS}" '
        f'xmlns:fo="{ODF_FO_NS}"'
    )


def _odf_styles_xml() -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles {_odf_content_namespaces()} office:version="{ODF_VERSION}">
  <office:styles>
    <style:style style:name="Standard" style:family="paragraph"/>
    <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph">
      <style:text-properties fo:font-weight="bold" fo:font-size="18pt"/>
    </style:style>
  </office:styles>
</office:document-styles>
"""


def _odf_meta_xml() -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="{ODF_OFFICE_NS}" office:version="{ODF_VERSION}">
  <office:meta/>
</office:document-meta>
"""


def _odf_settings_xml() -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings xmlns:office="{ODF_OFFICE_NS}" office:version="{ODF_VERSION}">
  <office:settings/>
</office:document-settings>
"""


def _odf_manifest_xml(media_type: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="{ODF_MANIFEST_NS}" manifest:version="{ODF_VERSION}">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="{media_type}"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"""


def _odt_paragraph(line: str, heading: bool = False) -> str:
    text = escape(str(line))
    if heading:
        return f'<text:h text:outline-level="1">{text}</text:h>'
    return f"<text:p>{text}</text:p>"


def _ods_row(row: list[Any]) -> str:
    cells = "".join(_ods_cell(value) for value in row)
    return f"<table:table-row>{cells}</table:table-row>"


def _ods_cell(value: Any) -> str:
    value = _xlsx_value(value)
    if value in (None, ""):
        return "<table:table-cell/>"
    if isinstance(value, bool):
        text = "TRUE" if value else "FALSE"
        return (
            f'<table:table-cell office:value-type="boolean" office:boolean-value="{str(value).lower()}">'
            f"<text:p>{text}</text:p></table:table-cell>"
        )
    if isinstance(value, (int, float)):
        return (
            f'<table:table-cell office:value-type="float" office:value="{value}">'
            f"<text:p>{value}</text:p></table:table-cell>"
        )
    text = escape(str(value))
    return f'<table:table-cell office:value-type="string"><text:p>{text}</text:p></table:table-cell>'


def _odp_page(slide: dict[str, Any], index: int) -> str:
    title = escape(str(slide.get("title") or f"Slide {index}"))
    bullets = [escape(str(item)) for item in slide.get("bullets") or []]
    bullet_items = "".join(f"<text:list-item><text:p>{bullet}</text:p></text:list-item>" for bullet in bullets)
    body = f"<text:list>{bullet_items}</text:list>" if bullet_items else "<text:p/>"
    return f"""<draw:page draw:name="Slide {index}" draw:master-page-name="Default">
  <draw:frame presentation:class="title" draw:name="Title {index}" svg:width="24cm" svg:height="2cm" svg:x="1.5cm" svg:y="1cm" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
    <draw:text-box><text:p>{title}</text:p></draw:text-box>
  </draw:frame>
  <draw:frame presentation:class="outline" draw:name="Content {index}" svg:width="24cm" svg:height="12cm" svg:x="1.5cm" svg:y="3.5cm" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
    <draw:text-box>{body}</draw:text-box>
  </draw:frame>
</draw:page>"""


def _docx(title: str, content: str) -> bytes:
    lines = [title] + [line for line in content.splitlines() if line.strip()]
    if len(lines) == 1:
        lines.append("")
    body = "".join(_docx_paragraph(line) for line in lines)
    return _zip_bytes({
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
        "word/document.xml": f"""<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>""",
    })


def _docx_paragraph(line: str) -> str:
    if not str(line).strip():
        return '<w:p><w:r><w:t xml:space="preserve">&#160;</w:t></w:r></w:p>'
    return f"<w:p><w:r><w:t>{escape(line)}</w:t></w:r></w:p>"


def _xlsx(title: str, content: str) -> bytes:
    rows = _xlsx_rows(title, content)
    sheet_rows = "".join(
        f'<row r="{row_idx}">{"".join(_xlsx_cell(row_idx, col_idx, value) for col_idx, value in enumerate(row, start=1))}</row>'
        for row_idx, row in enumerate(rows, start=1)
    )
    return _zip_bytes({
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""",
        "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""",
        "xl/worksheets/sheet1.xml": f"""<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{sheet_rows}</sheetData></worksheet>""",
    })


def _xlsx_rows(title: str, content: str) -> list[list[Any]]:
    parsed = _tabular_rows(content)
    if parsed:
        return parsed
    lines = [line for line in str(content or "").splitlines() if line.strip()]
    if lines:
        return [[title], *[[line] for line in lines]]
    return [[title]]


def _tabular_rows(content: str) -> list[list[Any]]:
    text = str(content or "").strip("\n")
    if not text.strip():
        return []
    lines = [line for line in text.splitlines() if line.strip()]
    markdown_rows = _markdown_table_rows(lines)
    if markdown_rows:
        return markdown_rows

    delimiter = "\t" if any("\t" in line for line in lines) else ("," if any("," in line for line in lines) else None)
    if not delimiter:
        return []
    return [[_xlsx_value(cell) for cell in row] for row in csv.reader(io.StringIO("\n".join(lines)), delimiter=delimiter)]


def _markdown_table_rows(lines: list[str]) -> list[list[Any]]:
    table_lines = [line.strip() for line in lines if line.strip().startswith("|") and line.strip().endswith("|")]
    if len(table_lines) < 2:
        return []
    rows = []
    for line in table_lines:
        cells = [cell.strip() for cell in line.strip("|").split("|")]
        if all(re.fullmatch(r":?-{3,}:?", cell or "") for cell in cells):
            continue
        rows.append([_xlsx_value(cell) for cell in cells])
    return rows


def _xlsx_cell(row_idx: int, col_idx: int, value: Any) -> str:
    ref = f"{_column_name(col_idx)}{row_idx}"
    value = _xlsx_value(value)
    if value in (None, ""):
        return f'<c r="{ref}"/>'
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"><v>{value}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'


def _xlsx_value(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    stripped = value.strip()
    if not stripped:
        return ""
    if stripped.lower() in {"true", "false"}:
        return stripped.lower() == "true"
    if re.fullmatch(r"[+-]?\d+", stripped) and not (len(stripped.lstrip("+-")) > 1 and stripped.lstrip("+-").startswith("0")):
        try:
            return int(stripped)
        except ValueError:
            return stripped
    if re.fullmatch(r"[+-]?(?:\d+\.\d*|\.\d+)(?:[eE][+-]?\d+)?", stripped) or re.fullmatch(r"[+-]?\d+[eE][+-]?\d+", stripped):
        try:
            return float(stripped)
        except ValueError:
            return stripped
    return stripped


def _column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _pptx(title: str, content: str) -> bytes:
    return pptx_writer.pptx_from_text(title, content)


_current = {}
def replace_document_bytes(file_id, data, **kwargs):
    doc = _current[file_id]
    Path(doc['path']).write_bytes(data)
    return doc
