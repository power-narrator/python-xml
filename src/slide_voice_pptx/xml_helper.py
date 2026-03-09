"""Shared XML helper utilities for PPTX manipulation."""

import xml.etree.ElementTree as ET

from .namespaces import NAMESPACE_CT, NSMAP_CT
from .xpath import (
    XPATH_CT_DEFAULT_BY_EXTENSION,
    XPATH_CT_OVERRIDE_BY_PATH_NAME,
)


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


def ensure_content_type_default(
    root: ET.Element, extension: str, content_type: str
) -> None:
    """Add a Default entry to [Content_Types].xml if not present.

    Args:
        root: Root element of [Content_Types].xml.
        extension: File extension.
        content_type: MIME content type.
    """
    if (
        root.find(
            XPATH_CT_DEFAULT_BY_EXTENSION.format(extension=extension),
            namespaces=NSMAP_CT,
        )
        is not None
    ):
        return

    ET.SubElement(
        root,
        f"{{{NAMESPACE_CT}}}Default",
        Extension=extension,
        ContentType=content_type,
    )


def ensure_content_type_override(
    root: ET.Element, path_name: str, content_type: str
) -> None:
    """Add an Override entry to [Content_Types].xml if not present.

    Args:
        root: Root element of [Content_Types].xml.
        path_name: Absolute package path (for example: /ppt/slides/slide1.xml).
        content_type: MIME content type for the path.
    """
    if (
        root.find(
            XPATH_CT_OVERRIDE_BY_PATH_NAME.format(path_name=path_name),
            namespaces=NSMAP_CT,
        )
    ) is not None:
        return

    ET.SubElement(
        root,
        f"{{{NAMESPACE_CT}}}Override",
        PartName=path_name,
        ContentType=content_type,
    )
