# Word Language Switching – Changelog

## 1.5.1
- **Changed gesture bindings:** NVDA 2026.2 reserves `NVDA+Shift+L` for the
  Magnifier, so the add-on's gestures were updated:
  - Toggle the add-on: `NVDA+Shift+L` → **`NVDA+Alt+W`**
  - Diagnostics script: `NVDA+Shift+K` → **`NVDA+Alt+K`**
  Both can still be reassigned via the Input Gestures dialog (category
  "Word language switching").
- **Updated to the current NVDA add-on template:**
  - Build tooling (`site_scons`, NVDATool + gettexttool) brought up to date
    with the latest template.
  - Added the `speechDictionaries` field to `buildVars.py` and `sconstruct`.
  - Adopted the template's development and CI files
    (`pyproject.toml`, `uv.lock`, `prek.toml`, `.python-version`, `.vscode`,
    `.github` workflows, `tests`, `docs`).
