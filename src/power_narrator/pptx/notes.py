"""Notes extraction and write support for PPTX slides."""

import xml.etree.ElementTree as ET
from pathlib import Path

from .content_types import ensure_content_type_overrides
from .exceptions import (
    InvalidPptxError,
    RelationshipIdNotFoundError,
    RelationshipTargetNotFoundError,
)
from .namespaces import (
    NAMESPACE_A,
    NAMESPACE_P,
    NAMESPACE_R,
    NAMESPACE_RELS,
    NSMAP,
    REL_TYPE_NOTES_MASTER,
    REL_TYPE_NOTES_SLIDE,
    REL_TYPE_THEME,
)
from .paths import (
    relative_target_path,
    rels_path_for_path,
    resolve_target_path,
    slide_rels_path,
)
from .rels import (
    add_relationship,
    find_relationship_target_by_type,
    get_relationship_id_target_map,
    read_rels_path,
)
from .xml_helper import ensure_child
from .xpath import (
    XPATH_NOTES_BODY_SHAPES,
    XPATH_NOTES_MASTER_ID_WITH_RID,
    XPATH_P_CNVPR_WITH_ID,
    XPATH_PARAGRAPH_TEXT,
    XPATH_SHAPE_PARAGRAPHS,
    XPATH_TXBODY_PARAGRAPHS,
)

CONTENT_TYPE_NOTES_MASTER = (
    "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
)
CONTENT_TYPE_NOTES_SLIDE = (
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
)
CONTENT_TYPE_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"

NOTES_MASTER_PATH = "ppt/notesMasters/notesMaster1.xml"
THEME2_PATH = "ppt/theme/theme2.xml"


def _extract_paragraphs(shape_element: ET.Element) -> list[str]:
    """Extract paragraph texts from a shape element.

    Args:
        shape_element: Shape XML element containing text body.

    Returns:
        List of paragraph strings.
    """
    paragraphs = []
    p_elements = shape_element.findall(XPATH_SHAPE_PARAGRAPHS, namespaces=NSMAP)

    for p_elem in p_elements:
        text_elements = p_elem.findall(XPATH_PARAGRAPH_TEXT, namespaces=NSMAP)
        para_text = "".join((t.text or "") for t in text_elements)
        paragraphs.append(para_text)

    return paragraphs


def extract_notes_text(notes_element: ET.Element) -> str:
    """Extract plain text from a notes slide XML element.

    Finds the body placeholder shape and extracts all text content.

    Args:
        notes_element: Parsed notesSlide XML element.

    Returns:
        Plain text content of the notes, with paragraphs joined by newlines.
    """
    paragraphs = []

    for shape in notes_element.findall(XPATH_NOTES_BODY_SHAPES, namespaces=NSMAP):
        paragraphs.extend(_extract_paragraphs(shape))

    return "\n".join(paragraphs)


def _next_shape_id(sp_tree: ET.Element) -> str:
    """Return the next available `p:cNvPr/@id` under a shape tree.

    Args:
        sp_tree: Parent shape tree element.

    Returns:
        String ID value greater than any existing `p:cNvPr/@id`.
    """
    return str(
        max(
            [
                int(shape_id)
                for c_nv_pr in sp_tree.findall(XPATH_P_CNVPR_WITH_ID, namespaces=NSMAP)
                if (shape_id := c_nv_pr.get("id")) and shape_id.isdigit()
            ]
            + [0]
        )
        + 1
    )


def _create_notes_body_placeholder_shape(
    sp_tree: ET.Element, shape_id: str
) -> ET.Element:
    """Create and return the notes body placeholder shape.

    Args:
        sp_tree: Parent shape tree element.
        shape_id: Value for the nested `p:cNvPr/@id` attribute.

    Returns:
        Created `p:sp` body-placeholder shape element.
    """
    body_shape = ET.SubElement(sp_tree, f"{{{NAMESPACE_P}}}sp")
    nv_sp_pr = ET.SubElement(body_shape, f"{{{NAMESPACE_P}}}nvSpPr")
    ET.SubElement(
        nv_sp_pr,
        f"{{{NAMESPACE_P}}}cNvPr",
        id=shape_id,
        name="Notes Placeholder 2",
    )
    ET.SubElement(nv_sp_pr, f"{{{NAMESPACE_P}}}cNvSpPr")
    nv_pr = ET.SubElement(nv_sp_pr, f"{{{NAMESPACE_P}}}nvPr")
    ET.SubElement(nv_pr, f"{{{NAMESPACE_P}}}ph", {"type": "body", "idx": "1"})
    ET.SubElement(body_shape, f"{{{NAMESPACE_P}}}spPr")
    ET.SubElement(body_shape, f"{{{NAMESPACE_P}}}txBody")

    return body_shape


def _ensure_notes_body_tx_body(notes_root: ET.Element) -> ET.Element:
    """Ensure and return the notes body placeholder `p:txBody` element.

    Args:
        notes_root: Parsed notes slide root element.

    Returns:
        Existing or newly created `p:txBody` element.
    """
    c_sld = ensure_child(notes_root, f"{{{NAMESPACE_P}}}cSld")
    sp_tree = ensure_child(c_sld, f"{{{NAMESPACE_P}}}spTree")
    body_shape = notes_root.find(XPATH_NOTES_BODY_SHAPES, namespaces=NSMAP)

    if body_shape is None:
        body_shape = _create_notes_body_placeholder_shape(
            sp_tree, shape_id=_next_shape_id(sp_tree)
        )

    return ensure_child(body_shape, f"{{{NAMESPACE_P}}}txBody")


def _set_notes_text(notes_root: ET.Element, text: str) -> None:
    """Replace body placeholder text with plain paragraph content.

    Args:
        notes_root: Parsed notes slide root element.
        text: Plain notes text where paragraphs are separated by newlines.
    """
    tx_body = _ensure_notes_body_tx_body(notes_root)

    for paragraph in tx_body.findall(XPATH_TXBODY_PARAGRAPHS, namespaces=NSMAP):
        tx_body.remove(paragraph)

    paragraphs = text.split("\n") if text else [""]

    for paragraph_text in paragraphs:
        paragraph = ET.SubElement(tx_body, f"{{{NAMESPACE_A}}}p")

        if paragraph_text:
            run = ET.SubElement(paragraph, f"{{{NAMESPACE_A}}}r")
            ET.SubElement(run, f"{{{NAMESPACE_A}}}t").text = paragraph_text


def _create_notes_slide_xml(text: str) -> ET.Element:
    """Create a notes slide XML root with placeholders and text body.

    Args:
        text: Plain notes text where paragraphs are separated by newlines.

    Returns:
        Notes slide root element.
    """
    notes_root = ET.Element(f"{{{NAMESPACE_P}}}notes")
    c_sld = ET.SubElement(notes_root, f"{{{NAMESPACE_P}}}cSld")
    sp_tree = ET.SubElement(c_sld, f"{{{NAMESPACE_P}}}spTree")

    nv_grp_sp_pr = ET.SubElement(sp_tree, f"{{{NAMESPACE_P}}}nvGrpSpPr")
    ET.SubElement(nv_grp_sp_pr, f"{{{NAMESPACE_P}}}cNvPr", id="1", name="")
    ET.SubElement(nv_grp_sp_pr, f"{{{NAMESPACE_P}}}cNvGrpSpPr")
    ET.SubElement(nv_grp_sp_pr, f"{{{NAMESPACE_P}}}nvPr")
    ET.SubElement(sp_tree, f"{{{NAMESPACE_P}}}grpSpPr")

    _create_notes_body_placeholder_shape(sp_tree, shape_id="3")
    tx_body = _ensure_notes_body_tx_body(notes_root)
    body_pr = ET.Element(f"{{{NAMESPACE_A}}}bodyPr")
    tx_body.insert(0, body_pr)

    _set_notes_text(notes_root, text)

    return notes_root


def _notes_filename_for_slide(slide_path: str) -> str:
    """Build notes filename for a slide path.

    Args:
        slide_path: Slide OOXML path.

    Returns:
        Notes slide filename.
    """
    slide_stem = Path(slide_path).stem
    suffix = slide_stem.replace("slide", "")
    return f"notesSlide{suffix}.xml"


def _default_clr_map() -> dict[str, str]:
    """Build default color mapping attributes for notes master.

    Returns:
        Attribute mapping for p:clrMap.
    """
    return {
        "bg1": "lt1",
        "tx1": "dk1",
        "bg2": "lt2",
        "tx2": "dk2",
        "accent1": "accent1",
        "accent2": "accent2",
        "accent3": "accent3",
        "accent4": "accent4",
        "accent5": "accent5",
        "accent6": "accent6",
        "hlink": "hlink",
        "folHlink": "folHlink",
    }


def _create_notes_master_xml() -> ET.Element:
    """Create a minimal notes master XML root.

    Returns:
        Notes master root element.
    """
    root = ET.Element(f"{{{NAMESPACE_P}}}notesMaster")
    c_sld = ET.SubElement(root, f"{{{NAMESPACE_P}}}cSld")
    sp_tree = ET.SubElement(c_sld, f"{{{NAMESPACE_P}}}spTree")

    nv_grp_sp_pr = ET.SubElement(sp_tree, f"{{{NAMESPACE_P}}}nvGrpSpPr")
    ET.SubElement(nv_grp_sp_pr, f"{{{NAMESPACE_P}}}cNvPr", id="1", name="")
    ET.SubElement(nv_grp_sp_pr, f"{{{NAMESPACE_P}}}cNvGrpSpPr")
    ET.SubElement(nv_grp_sp_pr, f"{{{NAMESPACE_P}}}nvPr")
    ET.SubElement(sp_tree, f"{{{NAMESPACE_P}}}grpSpPr")

    ET.SubElement(root, f"{{{NAMESPACE_P}}}clrMap", _default_clr_map())

    return root


def _create_theme2(work_dir: Path) -> None:
    """Create theme2 by cloning theme1.xml when needed.

    Args:
        work_dir: Extracted PPTX workspace root directory.

    Raises:
        InvalidPptxError: If `ppt/theme/theme1.xml` is missing.
    """
    theme2_path = work_dir / THEME2_PATH

    if theme2_path.exists():
        return

    theme1_rel_path = "ppt/theme/theme1.xml"
    theme1_path = work_dir / theme1_rel_path

    if not theme1_path.exists():
        raise InvalidPptxError(
            theme1_rel_path,
            "Required theme part is missing; cannot create theme2.xml",
        )

    theme2_path.parent.mkdir(parents=True, exist_ok=True)
    theme2_path.write_bytes(theme1_path.read_bytes())


def _find_theme_path_for_notes_master_rels(
    notes_master_path: str, notes_master_rels: ET.Element
) -> str | None:
    """Return the first theme path linked from notes master relationships.

    Args:
        notes_master_path: Notes-master path in the package.
        notes_master_rels: Parsed notes-master relationships XML element.

    Returns:
        Normalized theme path if found; otherwise `None`.
    """
    if theme_target := find_relationship_target_by_type(
        notes_master_rels,
        REL_TYPE_THEME,
    ):
        return resolve_target_path(notes_master_path, theme_target)

    return None


def _ensure_notes_master_files(work_dir: Path, notes_master_path: str) -> str:
    """Create default notes master file, theme rels, and theme if missing.

    Args:
        work_dir: Extracted PPTX workspace root directory.
        notes_master_path: Notes master path in package.

    Returns:
        Theme path for notes master.

    Raises:
        InvalidPptxError: If `ppt/theme/theme1.xml` is missing.
    """
    notes_master_file_path = work_dir / notes_master_path
    notes_master_rels_path = work_dir / rels_path_for_path(notes_master_path)
    notes_master_file_path.parent.mkdir(parents=True, exist_ok=True)
    notes_master_rels_path.parent.mkdir(parents=True, exist_ok=True)

    if not notes_master_file_path.exists():
        notes_master_root = _create_notes_master_xml()
        notes_master_file_path.write_bytes(
            ET.tostring(notes_master_root, encoding="UTF-8", xml_declaration=True)
        )

    if notes_master_rels_path.exists():
        notes_master_rels = ET.fromstring(notes_master_rels_path.read_bytes())
    else:
        notes_master_rels = ET.Element(f"{{{NAMESPACE_RELS}}}Relationships")

    theme_path = _find_theme_path_for_notes_master_rels(
        notes_master_path, notes_master_rels
    )

    if theme_path is None:
        _create_theme2(work_dir)
        theme_target = relative_target_path(notes_master_path, THEME2_PATH)
        add_relationship(notes_master_rels, REL_TYPE_THEME, theme_target)
        theme_path = THEME2_PATH
        ET.register_namespace("", NAMESPACE_RELS)
        notes_master_rels_path.write_bytes(
            ET.tostring(notes_master_rels, encoding="UTF-8", xml_declaration=True)
        )

    return theme_path


def _append_notes_master_id(presentation_root: ET.Element, rid: str) -> None:
    """Append a `p:notesMasterId` entry to the presentation root.

    Args:
        presentation_root: Parsed `p:presentation` root element.
        rid: Relationship ID to write to the `r:id` attribute.
    """
    notes_master_id_lst = presentation_root.find(f"{{{NAMESPACE_P}}}notesMasterIdLst")

    if notes_master_id_lst is None:
        notes_master_id_lst = ET.Element(f"{{{NAMESPACE_P}}}notesMasterIdLst")
        sld_master_id_lst = presentation_root.find(f"{{{NAMESPACE_P}}}sldMasterIdLst")

        if sld_master_id_lst is None:
            presentation_root.insert(1, notes_master_id_lst)
        else:
            children = list(presentation_root)
            presentation_root.insert(
                children.index(sld_master_id_lst) + 1, notes_master_id_lst
            )

    notes_master_id = ET.SubElement(
        notes_master_id_lst, f"{{{NAMESPACE_P}}}notesMasterId"
    )
    notes_master_id.set(f"{{{NAMESPACE_R}}}id", rid)


def _ensure_notes_master(work_dir: Path) -> str:
    """Ensure notes master files and relationships exist.

    Args:
        work_dir: Extracted PPTX workspace root directory.

    Returns:
        Notes master package path.

    Raises:
        InvalidPptxError: If `ppt/theme/theme1.xml` is missing.
        RelsNotFoundError: If the presentation relationships file is missing.
        RelationshipIdNotFoundError: If `notesMasterId/@r:id` has no matching
            relationship entry in presentation relationships.
    """
    presentation_path = work_dir / "ppt/presentation.xml"
    presentation_root = ET.fromstring(presentation_path.read_bytes())
    presentation_rels_path = work_dir / "ppt/_rels/presentation.xml.rels"
    presentation_rels = read_rels_path(presentation_rels_path)
    notes_master_rels = get_relationship_id_target_map(
        presentation_rels,
        rel_type=REL_TYPE_NOTES_MASTER,
    )

    notes_master_id = presentation_root.find(
        XPATH_NOTES_MASTER_ID_WITH_RID, namespaces=NSMAP
    )

    if notes_master_id is not None:
        notes_master_rid = notes_master_id.get(f"{{{NAMESPACE_R}}}id", "")
        notes_master_target = notes_master_rels.get(notes_master_rid)

        if notes_master_target is None:
            raise RelationshipIdNotFoundError(
                "ppt/_rels/presentation.xml.rels", notes_master_rid
            )
    else:
        if notes_master_rels:
            notes_master_rid, notes_master_target = next(
                iter(notes_master_rels.items())
            )
        else:
            notes_master_target = relative_target_path(
                "ppt/presentation.xml", NOTES_MASTER_PATH
            )
            notes_master_rid = add_relationship(
                presentation_rels, REL_TYPE_NOTES_MASTER, notes_master_target
            )
            ET.register_namespace("", NAMESPACE_RELS)
            presentation_rels_path.write_bytes(
                ET.tostring(presentation_rels, encoding="UTF-8", xml_declaration=True)
            )

        _append_notes_master_id(presentation_root, notes_master_rid)
        presentation_path.write_bytes(
            ET.tostring(presentation_root, encoding="UTF-8", xml_declaration=True)
        )

    notes_master_path = resolve_target_path("ppt/presentation.xml", notes_master_target)
    theme_path = _ensure_notes_master_files(work_dir, notes_master_path)

    ensure_content_type_overrides(
        work_dir,
        {
            (f"/{notes_master_path}", CONTENT_TYPE_NOTES_MASTER),
            (f"/{theme_path}", CONTENT_TYPE_THEME),
        },
    )

    return notes_master_path


def write_slide_notes(work_dir: Path, slide_path: str, text: str) -> None:
    """Write notes text for a slide, creating notes files if needed.

    Args:
        work_dir: Extracted PPTX workspace root directory.
        slide_path: OOXML path (for example: ppt/slides/slide1.xml).
        text: Plain notes text where paragraphs are separated by newlines.

    Raises:
        RelationshipTargetNotFoundError: If a relationship target path resolves
            to a missing notes XML file.
        InvalidPptxError: If `ppt/theme/theme1.xml` is missing.
        RelationshipIdNotFoundError: If `notesMasterId/@r:id` has no matching
            relationship entry in presentation relationships.
        RelsNotFoundError: If a required slide, presentation, or notes-master
            relationship file is missing.
    """
    ET.register_namespace("a", NAMESPACE_A)
    ET.register_namespace("p", NAMESPACE_P)
    ET.register_namespace("r", NAMESPACE_R)

    rels_path = slide_rels_path(work_dir, slide_path)
    slide_rels = read_rels_path(rels_path)

    if notes_target := find_relationship_target_by_type(
        slide_rels, REL_TYPE_NOTES_SLIDE
    ):
        notes_xml_path = resolve_target_path(slide_path, notes_target)
        notes_path = work_dir / notes_xml_path

        if not notes_path.exists():
            raise RelationshipTargetNotFoundError(
                rels_path_for_path(slide_path), notes_xml_path
            )

        notes_root = ET.fromstring(notes_path.read_bytes())
        _set_notes_text(notes_root, text)
        notes_path.write_bytes(
            ET.tostring(notes_root, encoding="UTF-8", xml_declaration=True)
        )
        return

    notes_master_path = _ensure_notes_master(work_dir)
    notes_filename = _notes_filename_for_slide(slide_path)
    notes_xml_path = f"ppt/notesSlides/{notes_filename}"
    notes_rels_path = work_dir / "ppt/notesSlides/_rels" / f"{notes_filename}.rels"
    notes_path = work_dir / notes_xml_path

    notes_path.parent.mkdir(parents=True, exist_ok=True)
    notes_rels_path.parent.mkdir(parents=True, exist_ok=True)

    notes_root = _create_notes_slide_xml(text)
    notes_path.write_bytes(
        ET.tostring(notes_root, encoding="UTF-8", xml_declaration=True)
    )

    notes_rels_root = ET.Element(f"{{{NAMESPACE_RELS}}}Relationships")
    notes_master_target = relative_target_path(notes_xml_path, notes_master_path)
    add_relationship(notes_rels_root, REL_TYPE_NOTES_MASTER, notes_master_target)
    slide_target = relative_target_path(notes_xml_path, slide_path)
    add_relationship(
        notes_rels_root,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        slide_target,
    )
    ET.register_namespace("", NAMESPACE_RELS)
    notes_rels_path.write_bytes(
        ET.tostring(notes_rels_root, encoding="UTF-8", xml_declaration=True)
    )

    notes_target_from_slide = relative_target_path(slide_path, notes_xml_path)
    add_relationship(slide_rels, REL_TYPE_NOTES_SLIDE, notes_target_from_slide)
    rels_path.write_bytes(
        ET.tostring(slide_rels, encoding="UTF-8", xml_declaration=True)
    )

    ensure_content_type_overrides(
        work_dir,
        {
            (f"/ppt/notesSlides/{notes_filename}", CONTENT_TYPE_NOTES_SLIDE),
            (f"/{notes_master_path}", CONTENT_TYPE_NOTES_MASTER),
        },
    )
