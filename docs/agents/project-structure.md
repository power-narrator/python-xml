# Project Structure

```
slide-voice-app/
├── src/slide_voice_app/     # GUI application source code
│   ├── __main__.py          # GUI application entry point
│   ├── tts/                 # Text-to-speech providers and audio generation logic
│   ├── ui/                  # QML files for UI
│   ├── qml_modules/         # QML custom modules outputs (qmldir, qml, qmltypes)
│   └── rc_resources.py      # Generated resource file (do not edit manually)
├── src/slide_voice_pptx/    # PPTX library with CLI entry point
├── scripts/                 # Build and development scripts
│   ├── run.py               # Run application in development mode
│   ├── build.py             # Build application for distribution
│   └── utils.py             # Shared utilities for scripts
├── docs/                    # Documentation
│   ├── dev/                 # Developer documentation
│   └── user/                # User documentation
├── resources.qrc            # Qt resource collection file
└── pyproject.toml           # Project configuration
```
