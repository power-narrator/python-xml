"""Path helpers for PPTX OOXML package paths and relationships."""

import posixpath
from pathlib import Path


def resolve_target_path(source_path: str, target: str) -> str:
    """Resolve a relationship target to a normalized package path.

    Args:
        source_path: Path that owns the relationship.
        target: Relationship target value.

    Returns:
        Normalized package path for the target.
    """
    source_dir = posixpath.dirname(source_path)
    return posixpath.normpath(posixpath.join(source_dir, target))


def relative_target_path(source_path: str, target_path: str) -> str:
    """Build a relationship target from one package path to another.

    Args:
        source_path: Path that will own the relationship.
        target_path: Destination package path.

    Returns:
        Relative relationship target path.
    """
    source_dir = posixpath.dirname(source_path)
    destination_path = target_path
    return posixpath.relpath(destination_path, source_dir or ".")


def rels_path_for_path(path_value: str) -> str:
    """Build relationship-file path for a package path.

    Args:
        path_value: Package path.

    Returns:
        Relationship XML path under the source `_rels` directory.
    """
    parent_dir = posixpath.dirname(path_value)
    file_name = posixpath.basename(path_value)
    return posixpath.join(parent_dir, "_rels", f"{file_name}.rels")


def slide_rels_path(work_dir: Path, slide_path: str) -> Path:
    """Build absolute path to a slide relationship file in workspace.

    Args:
        work_dir: Extracted PPTX workspace root directory.
        slide_path: Slide OOXML path.

    Returns:
        Absolute path to the slide `.rels` file.
    """
    return work_dir / rels_path_for_path(slide_path)
