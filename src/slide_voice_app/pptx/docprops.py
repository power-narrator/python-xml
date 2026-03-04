"""Helpers for updating PPTX docProps metadata."""

import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

from .namespaces import NAMESPACE_DCTERMS, NAMESPACE_XSI, REL_TYPE_NOTES_SLIDE
from .rels import find_relationship_target_by_type


def update_core_xml_modified(core_content: bytes) -> bytes:
    """Update dcterms:modified timestamp in core.xml.

    Args:
        core_content: Raw bytes of docProps/core.xml.

    Returns:
        Updated XML bytes.
    """
    ET.register_namespace("dcterms", NAMESPACE_DCTERMS)
    ET.register_namespace("xsi", NAMESPACE_XSI)
    root = ET.fromstring(core_content)
    modified = root.find(f"{{{NAMESPACE_DCTERMS}}}modified")

    if modified is None:
        modified = ET.SubElement(root, f"{{{NAMESPACE_DCTERMS}}}modified")

    now = datetime.now(timezone.utc).replace(microsecond=0)
    modified.text = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    modified.set(f"{{{NAMESPACE_XSI}}}type", "dcterms:W3CDTF")

    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def update_app_xml_notes_count(app_content: bytes, notes_count: int) -> bytes:
    """Update Notes count in docProps/app.xml.

    Args:
        app_content: Raw bytes of docProps/app.xml.
        notes_count: Number of slides with notes slides.

    Returns:
        Updated XML bytes.
    """
    namespace = (
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    )
    root = ET.fromstring(app_content)
    notes = root.find(f"{{{namespace}}}Notes")

    if notes is None:
        notes = ET.SubElement(root, f"{{{namespace}}}Notes")

    notes.text = str(notes_count)

    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def count_slides_with_notes(work_dir: Path) -> int:
    """Count slides that reference a notesSlide relationship.

    Args:
        work_dir: Extracted PPTX workspace root.

    Returns:
        Number of slides with notes relationships.
    """
    rels_dir = work_dir / "ppt/slides/_rels"

    if not rels_dir.exists():
        return 0

    count = 0

    for rels_path in rels_dir.glob("slide*.xml.rels"):
        rels_root = ET.fromstring(rels_path.read_bytes())

        if (
            find_relationship_target_by_type(rels_root, REL_TYPE_NOTES_SLIDE)
            is not None
        ):
            count += 1

    return count
