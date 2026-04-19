import subprocess
import sys

from utils import (
    compile_resources,
    generate_qml_module_artifacts,
)


def run_dev():
    """Runs the application."""
    try:
        generate_qml_module_artifacts()
        compile_resources()
        _ = subprocess.run(["uv", "run", "-m", "power_narrator"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Running application failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    run_dev()
