import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
UI_PKG_DIR = BASE_DIR / "src" / "power_narrator" / "ui"


@dataclass(frozen=True)
class QmlModuleSpec:
    """Configuration for generating QML module artifacts."""

    import_name: str
    major_version: int
    minor_version: int
    source_files: list[str]


QML_MODULE_SPECS: list[QmlModuleSpec] = [
    QmlModuleSpec(
        import_name="PowerNarrator",
        major_version=1,
        minor_version=0,
        source_files=["models.py", "tts_manager.py", "pptx_manager.py"],
    ),
]


def compile_resources() -> None:
    """Runs the pyside6-rcc compiler via uv."""
    try:
        _ = subprocess.run(
            [
                "uv",
                "run",
                "pyside6-rcc",
                str(UI_PKG_DIR / "resources.qrc"),
                "-o",
                str(UI_PKG_DIR / "rc_resources.py"),
            ],
            check=True,
        )
    except subprocess.CalledProcessError as e:
        print(f"Error compiling resources: {e}")
        sys.exit(1)


def _generate_qml_module_artifacts(spec: QmlModuleSpec) -> None:
    module_dir = UI_PKG_DIR / "qml_modules" / spec.import_name

    for source_file in spec.source_files:
        if not (module_dir / source_file).exists():
            raise FileNotFoundError(f"QML module source file not found: {source_file}")

    qmldir_path = module_dir / "qmldir"
    metatypes_path = module_dir / "modulemetatypes.json"
    qmltypes_path = module_dir / "module.qmltypes"
    registrations_path = module_dir / "module_qmltyperegistrations.cpp"

    qmldir_path.write_text(f"module {spec.import_name}\ntypeinfo module.qmltypes")

    try:
        _ = subprocess.run(
            [
                "uv",
                "run",
                "pyside6-metaobjectdump",
                "-o",
                metatypes_path,
                *[(module_dir / path) for path in spec.source_files],
            ],
            check=True,
        )
    except subprocess.CalledProcessError as e:
        print(f"Error generating metatypes for {spec.import_name}: {e}")
        sys.exit(1)

    try:
        _ = subprocess.run(
            [
                "uv",
                "run",
                "pyside6-qmltyperegistrar",
                "--generate-qmltypes",
                str(qmltypes_path),
                "-o",
                str(registrations_path),
                str(metatypes_path),
                "--import-name",
                spec.import_name,
                "--major-version",
                str(spec.major_version),
                "--minor-version",
                str(spec.minor_version),
            ],
            check=True,
        )
    except subprocess.CalledProcessError as e:
        print(f"Error generating qmltypes for {spec.import_name}: {e}")
        sys.exit(1)


def generate_qml_module_artifacts() -> None:
    """Generate qmltypes and registration artifacts for QML modules."""
    for spec in QML_MODULE_SPECS:
        _generate_qml_module_artifacts(spec)


if __name__ == "__main__":
    generate_qml_module_artifacts()
    compile_resources()
