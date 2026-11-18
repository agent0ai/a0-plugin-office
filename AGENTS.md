# Office community plugin

## Purpose

Restore Agent Zero's ODF-first Writer, Calc and Impress workflows as a standalone community plugin named `office`.

## Ownership

- `helpers/document_store.py` preserves scoped paths, SQLite metadata, atomic writes and version backups under `usr/office`.
- `helpers/runtime.py` owns lazy sandbox lifecycle, private RPC, document synchronization and VNC registration.
- `sandbox/` owns LibreOffice/UNO and format dependencies, first-use setup and native document operations.
- `api/`, `tools/`, `skills/` and `webui/` expose the feature; `hooks.py` owns registration and cleanup.

## Local Contracts

- `plugin.yaml` declares `a0_version: ">=3.0"`; keep the minimum supported Agent Zero version explicit.
- Installed Python imports use `usr.plugins.office`. Keep plugin files at this repository root.
- Use Agent Zero's sandbox manager, explicit read-only code/Tuxee mounts and child lifetime. Do not mount the entire usr directory or install packages in the outer environment.
- Mount only `sandbox/` as code; granting mount ACLs on extension files triggers the plugin reloader. RPC uses a private directory under `/run`, mounted into the guest; clear it after stopping but retain the directory for mount validation until uninstall.
- Reuse the Tuxee Xvnc bundle and bundled noVNC code. No Xpra, Xfce, Collabora or alternate desktop service.
- Preserve authentication, same-origin and CSRF checks. VNC passwords stay session-scoped; only loopback/private socket listeners are permitted.
- Setup is automatic on first use, serialized and retryable, with visible progress. UI startup must stay responsive during installation.
- Preserve the original document on save/export/rename failure and never overwrite conflicting external edits. Keep user documents and version history on uninstall.
- Keep the Office toolbar on one row: File owns open/save/export/close, Edit owns rename/history/discard, Office creates Writer/Calc/Impress documents. Menus retain busy/active-document guards and keyboard dismissal.
- Ordinary tool results do not open UI unless requested. Text/Markdown remain in Agent Zero's Editor.

## Work Guidance

Reuse the prior document-format algorithms and current host APIs; isolate optional dependencies in the guest. Keep legacy parity and actual limitations in README.

## Verification

Run focused unit tests in the framework Python and live Docker checks for cold/warm setup, Writer/Calc/Impress editing, conversion, versions, conflicts, VNC input, resize and lifecycle. Capture screenshots without private data.

`tests/live_backend.py` and `tests/live_ui.py` are explicit disposable-container checks, not default pytest collection. The UI check needs test-only Playwright and a running A0 WebUI at its discovered URL. The backend check owns a separate runtime; stop Office in the WebUI first.

## Child DOX Index

No child DOX files.
