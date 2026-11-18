"""Saved format operations run in the Office guest, with host-owned version history."""
import base64
from usr.plugins.office.helpers import document_store
from usr.plugins.office.helpers.runtime import get_runtime


def read_artifact(doc, max_chars=12000):
    runtime = get_runtime()
    with runtime.lock:
        doc = runtime.sync(doc['file_id']).get('document', doc)
        return runtime.call('read', **runtime.payload(doc), max_chars=max_chars)


def edit_artifact(doc, **kwargs):
    runtime = get_runtime()
    with runtime.lock:
        doc = runtime.sync(doc['file_id']).get('document', doc)
        result = runtime.call('edit', **runtime.payload(doc), edit=kwargs)
        updated = document_store.replace_document_bytes(doc['file_id'],base64.b64decode(result['data'],validate=True),actor='office_artifact:edit')
        runtime.edited(updated)
        return updated, result['details']
