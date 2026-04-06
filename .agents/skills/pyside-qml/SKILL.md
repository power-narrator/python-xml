---
name: pyside-qml
description: PySide6 and QML workflow guidance for this repository. Use when Codex is adding or updating QML files, maintaining Qt resource declarations, regenerating QML module artifacts, or adjusting project wiring for PySide6 and QML modules.
---

# PySide6 / QML

Follow these rules for PySide6 and QML changes in this repository.

## Run The QML Utility

Use this command to manually compile Qt resources and QML module artifacts:

```bash
uv run scripts/utils.py
```

Default build and run scripts already perform this step automatically.

## Update Project Wiring

Add custom modules to `scripts/utils.py` so type generation includes them.

Add every new QML file to `resources.qrc` with the correct alias.
