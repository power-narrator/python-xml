import subprocess
import sys

from utils import (
    BASE_DIR,
    PKG_DIR,
    compile_resources,
    generate_qml_module_artifacts,
)


def run_build():
    """Builds the application using Nuitka."""
    generate_qml_module_artifacts()
    compile_resources()

    try:
        _ = subprocess.run(
            [
                "uv",
                "run",
                "python",
                "-m",
                "nuitka",
                "--enable-plugin=pyside6",
                "--include-qt-plugins=qml,multimedia",
                "--output-filename=slide-voice-app",
                "--mode=app",
                f"--output-dir={BASE_DIR / 'dist'}",
                "--include-data-files=src/slide_voice_pptx/resources/narration-icon.png=slide_voice_pptx/resources/narration-icon.png",
                PKG_DIR,
            ],
            check=True,
        )
    except subprocess.CalledProcessError as e:
        print(f"Nuitka build failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    run_build()
