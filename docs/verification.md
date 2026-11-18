# Verification

Tested in a disposable Linux x86_64 Docker container built from Agent Zero v3 commit `d87bd6d0`, using its global Python 3.12 interpreter, PRoot overlay sandbox manager and shared Tuxee display bundle. The container used Docker's normal seccomp policy with `no-new-privileges`, without privileged mode or extra capabilities. Office packages installed only inside `service-office`; Playwright/Chromium were test-only dependencies in the disposable container.

- **13 regression tests:** version restoration, private backups, preserved file permissions, failed-write recovery, external-edit conflicts, Save As collisions/remapping, scoped paths and symlinks, viewer ownership/origin/authentication checks, retry state, reusable RPC mount cleanup, and correct exported-file metadata.
- **Real backend:** create/read/edit/open/save/restore/close for ODT, DOCX, ODS, XLSX, ODP and PPTX; spreadsheet formulas and an embedded XLSX column chart; PDF export and paired ODF ↔ OOXML conversion for all formats; tool dispatch, rename, host Save As and native UNO Save As.
- **Real WebUI:** Writer, Calc and Impress through noVNC; keyboard and clipboard input saved into the document; native window switching/closing; floating-to-canvas docking, display resize and reconnect; Office with Browser disabled; disable/re-enable and retained installation.
- **Lifecycle and admission:** normal management API uninstall removes the plugin and its sandbox while preserving documents/history; ZIP installation runs the hook and discovers the API without eagerly provisioning packages. Live viewer requests reject cross-origin access, and API requests reject invalid CSRF tokens.
- **Setup:** successful first-use package installation, concurrent startup calls with responsive status, and warm sandbox reuse. Installed private sandbox disk measured approximately **588 MiB** including test files; warm service restart measured **1.5–2.1 seconds** on this VPS. These are observations, not performance guarantees, and exclude the shared base image.

Reproduce in a disposable A0 container with this repository installed at `/a0/usr/plugins/office`:

```sh
cd /a0
python -m pytest usr/plugins/office/tests -q
# Stop Office in the WebUI before the backend check owns its runtime:
python -m usr.plugins.office.tests.live_backend
# Start A0, discover its URL, then supply it explicitly:
python -m usr.plugins.office.tests.live_ui "$A0_TEST_URL" /tmp/office-shots
```

`tests/live_lifecycle.py` tests uninstall and ZIP installation through the normal management API; it deliberately removes the container's installed Office plugin. Copy it outside the plugin before running. All live scripts create disposable test documents.

The compact File/Edit/Office toolbar was checked separately in Chromium with the real Alpine templates and stubbed action callbacks: narrow/wide layouts, every menu action, export formats, disabled states, menu switching, Escape/focus return, outside clicks and window blur. [Toolbar preview](toolbar.png). The full Docker suite was not rerun for this layout-only change.

Earlier runtime screenshots (before the toolbar update): [Writer](writer.png), [Calc](calc.png), [Impress](impress.png), [docked Office](docked.png).

Not verified: ARM64/native-host operation, a model-driven agent conversation, every LibreOffice dialog or complex Office-format fidelity. The old document-format algorithms retain their existing limitations (for example, direct ODS chart editing and formatting loss during structural rewrites). Existing files can be opened, but retired built-in `_office` metadata/version-history migration is not implemented. Force-killing the container can lose unsaved native edits. Publishing this repository and listing it in the Plugin Hub are separate steps.
