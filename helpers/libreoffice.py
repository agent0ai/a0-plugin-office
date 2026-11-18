from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any


ODF_MIMETYPES = {
    "odt": "application/vnd.oasis.opendocument.text",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "odp": "application/vnd.oasis.opendocument.presentation",
}


def collect_status():
    from usr.plugins.office.helpers.runtime import get_runtime
    return {'ok':True, **get_runtime().status()}


def validate_docx(path: str | Path) -> dict[str, Any]:
    source = Path(path)
    if not source.exists():
        return {"ok": False, "error": f"File not found: {source}"}
    try:
        with zipfile.ZipFile(source) as archive:
            archive.getinfo("[Content_Types].xml")
            archive.getinfo("word/document.xml")
    except Exception as exc:
        return {"ok": False, "error": f"DOCX package validation failed: {exc}"}

    return {'ok': True}


def validate_odf(path: str | Path) -> dict[str, Any]:
    source = Path(path)
    if not source.exists():
        return {"ok": False, "error": f"File not found: {source}"}
    ext = source.suffix.lower().lstrip(".")
    expected_mimetype = ODF_MIMETYPES.get(ext)
    if not expected_mimetype:
        return {"ok": False, "error": f"Unsupported ODF extension: {ext}"}
    try:
        with zipfile.ZipFile(source) as archive:
            first = archive.infolist()[0]
            mimetype = archive.read("mimetype").decode("utf-8")
            archive.getinfo("content.xml")
            archive.getinfo("META-INF/manifest.xml")
    except Exception as exc:
        return {"ok": False, "error": f"ODF package validation failed: {exc}"}
    if first.filename != "mimetype" or first.compress_type != zipfile.ZIP_STORED:
        return {"ok": False, "error": "ODF mimetype must be the first uncompressed package entry."}
    if mimetype != expected_mimetype:
        return {"ok": False, "error": f"ODF mimetype mismatch: expected {expected_mimetype}, got {mimetype}"}
    return {"ok": True}


def convert_document(path, target_format, output_dir=None):
    import base64
    from usr.plugins.office.helpers import document_store
    from usr.plugins.office.helpers.runtime import get_runtime
    source = Path(path)
    target_format = str(target_format).lower().lstrip('.')
    if target_format not in {'odt','ods','odp','docx','xlsx','pptx','pdf'}:
        raise ValueError('Unsupported export format')
    target = (Path(output_dir) if output_dir else source.parent) / (source.stem + '.' + target_format)
    target = document_store.normalize_path(target)
    if target.exists():
        raise FileExistsError(f'Export target already exists: {target}')
    doc = document_store.register_document(source)
    runtime = get_runtime()
    with runtime.lock:
        doc = runtime.sync(doc['file_id']).get('document', doc)
        result = runtime.call('export', **runtime.payload(doc), target_format=target_format)
        document_store._write_atomic(target, base64.b64decode(result['data'],validate=True), source=source)
    return {'ok':True, 'path':str(target)}
