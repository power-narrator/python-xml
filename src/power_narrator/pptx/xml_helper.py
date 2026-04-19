"""Shared XML helper utilities for PPTX manipulation."""

import xml.etree.ElementTree as ET


def ensure_child(
    parent: ET.Element, tag: str, attrs: dict[str, str] | None = None
) -> ET.Element:
    """Find or create a child element with matching attributes.

    Searches for the first child element with the given tag where all
    specified attributes match. If no match is found, creates a new
    child element with the given tag and attributes.

    Args:
        parent: Parent element to search under.
        tag: Fully-qualified tag name.
        attrs: Attributes to match and use when creating a new element.

    Returns:
        The existing or newly created child element.
    """
    if not attrs:
        child = parent.find(tag)

        if child is not None:
            return child
    else:
        predicates = "".join(f"[@{key}='{value}']" for key, value in attrs.items())
        child = parent.find(f"{tag}{predicates}")

        if child is not None:
            return child

    return ET.SubElement(parent, tag, attrs or {})
