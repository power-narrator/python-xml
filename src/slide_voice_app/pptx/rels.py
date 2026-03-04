"""Relationship (.rels) file management for PPTX."""

import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from .exceptions import RelsNotFoundError
from .namespaces import NAMESPACE_RELS, NSMAP_RELS
from .xpath import (
    XPATH_RELATIONSHIP_WITH_ID,
)


def get_relationship_id_target_map(
    rels_element: ET.Element,
    rel_type: str | None = None,
    only_ids: set[str] | None = None,
) -> dict[str, str]:
    """Build a relationship ID -> target map with optional filters.

    Args:
        rels_element: Parsed .rels XML element.
        rel_type: Optional relationship type URI filter.
        only_ids: Optional set of relationship IDs to include.

    Returns:
        Mapping of relationship ID to target path.
    """
    relationship_map: dict[str, str] = {}

    for rel in rels_element.findall(XPATH_RELATIONSHIP_WITH_ID, namespaces=NSMAP_RELS):
        rid = rel.get("Id")
        target = rel.get("Target")

        if rid is None or target is None:
            continue

        if rel_type is not None and rel.get("Type") != rel_type:
            continue

        if only_ids is not None and rid not in only_ids:
            continue

        relationship_map[rid] = target

    return relationship_map


def read_rels(zip_file: ZipFile, rels_path: str) -> ET.Element:
    """Read and parse a relationship (.rels) file.

    Args:
        zip_file: Open ZipFile instance.
        rels_path: Path to the .rels file.

    Returns:
        Parsed XML Element of the relationships.

    Raises:
        RelsNotFoundError: If the .rels file does not exist in the archive.
    """
    try:
        content = zip_file.read(rels_path)
    except Exception as e:
        raise RelsNotFoundError(rels_path) from e

    return ET.fromstring(content)


def read_rels_path(rels_path: Path) -> ET.Element:
    """Read and parse a relationship (.rels) file from disk.

    Args:
        rels_path: Path to the .rels file.

    Returns:
        Parsed XML Element of the relationships.

    Raises:
        RelsNotFoundError: If the .rels file does not exist on disk.
    """
    if not rels_path.exists():
        raise RelsNotFoundError(str(rels_path))

    return ET.fromstring(rels_path.read_bytes())


def get_relationships_target_by_type(
    rels_element: ET.Element,
    rel_type: str,
) -> list[str]:
    """Find all relationships matching a specific type.

    Args:
        rels_element: Parsed .rels XML element.
        rel_type: Relationship type URI to find.

    Returns:
        List of matching relationship with targets.
    """
    return list(
        get_relationship_id_target_map(rels_element, rel_type=rel_type).values()
    )


def find_relationship_target_by_type(
    rels_element: ET.Element,
    rel_type: str,
) -> str | None:
    """Find any relationship target by type, return target or None.

    Args:
        rels_element: Parsed .rels XML element.
        rel_type: Relationship type URI.

    Returns:
        A matching target path if found, None otherwise.
    """
    return next(
        iter(get_relationship_id_target_map(rels_element, rel_type=rel_type).values()),
        None,
    )


def find_relationship_by_type_and_target(
    rels_element: ET.Element,
    rel_type: str,
    target: str,
) -> str | None:
    """Find existing relationship with matching type and target, return rId or None.

    Args:
        rels_element: Parsed .rels XML element.
        rel_type: Relationship type URI.
        target: Target path for the relationship.

    Returns:
        Relationship ID (rId) if found, None otherwise.
    """
    for rid, rel_target in get_relationship_id_target_map(
        rels_element,
        rel_type=rel_type,
    ).items():
        if rel_target == target:
            return rid

    return None


def get_next_rid(rels_element: ET.Element) -> str:
    """Get the next available relationship ID.

    Args:
        rels_element: Parsed .rels XML element.

    Returns:
        Next available rId.
    """
    ids = [
        int(rid)
        for rel in rels_element.findall(
            XPATH_RELATIONSHIP_WITH_ID,
            namespaces=NSMAP_RELS,
        )
        if (id := rel.get("Id")) and id.startswith("rId") and (rid := id[3:]).isdigit()
    ]
    return f"rId{max(ids, default=0) + 1}"


def add_relationship(
    rels_element: ET.Element,
    rel_type: str,
    target: str,
    rid: str | None = None,
) -> str:
    """Add a new relationship to a .rels element.

    Args:
        rels_element: Parsed .rels XML element to modify.
        rel_type: Relationship type URI.
        target: Target path (relative).
        rid: Optional specific rId to use; auto-generated if None.

    Returns:
        The relationship ID used.
    """
    if rid is None:
        rid = get_next_rid(rels_element)

    rel = ET.SubElement(
        rels_element,
        f"{{{NAMESPACE_RELS}}}Relationship",
    )
    rel.set("Id", rid)
    rel.set("Type", rel_type)
    rel.set("Target", target)

    return rid
