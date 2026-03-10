"""Insert audio into PPTX slides with autoplay on slide start."""

import hashlib
import re
import uuid
import xml.etree.ElementTree as ET
from importlib.resources import files
from pathlib import Path

from ..exceptions import SlideXmlNotFoundError
from ..namespaces import (
    NAMESPACE_A,
    NAMESPACE_A16,
    NAMESPACE_CT,
    NAMESPACE_P,
    NAMESPACE_P14,
    NAMESPACE_R,
    NAMESPACE_RELS,
    REL_TYPE_AUDIO,
    REL_TYPE_IMAGE,
    REL_TYPE_MEDIA,
)
from ..paths import slide_rels_path
from ..rels import add_relationship, find_relationship_by_type_and_target
from ..xml_helper import ensure_content_type_default
from .audio_timing import (
    compute_next_delay,
    create_audio_node,
    create_command_node,
    get_next_shape_id,
    get_next_timing_id,
    get_or_create_audio_parent,
    get_or_create_command_parent,
    get_or_create_pic_parent,
)

DEFAULT_ICON_X = 12479915
DEFAULT_ICON_Y = -126134
DEFAULT_ICON_CX = 812800
DEFAULT_ICON_CY = 812800

AUDIO_ICON_BYTES = (
    files("slide_voice_pptx").joinpath("resources", "narration-icon.png").read_bytes()
)
AUDIO_ICON_HASH = hashlib.sha256(AUDIO_ICON_BYTES).hexdigest()


def _find_media_files(media_dir: Path, prefix: str, ext: str) -> list[str]:
    """Find existing media files matching prefix and extension.

    Args:
        media_dir: Path to ppt/media directory.
        prefix: Filename prefix.
        ext: File extension without dot.

    Returns:
        List of filenames, not full paths.
    """
    pattern = re.compile(rf"^{prefix}(\d+)\.{ext}$")

    return [
        file_path.name
        for file_path in media_dir.iterdir()
        if file_path.is_file() and pattern.match(file_path.name)
    ]


def _next_media_filename(existing: list[str], prefix: str, ext: str) -> str:
    """Allocate next available media filename.

    Args:
        existing: List of existing media filenames.
        prefix: Filename prefix.
        ext: File extension without dot.

    Returns:
        Next available filename like 'media1.mp3'.
    """
    pattern = re.compile(rf"^{prefix}(\d+)\.{ext}$")
    max_num = max(
        [int(match.group(1)) for name in existing if (match := pattern.match(name))],
        default=0,
    )

    return f"{prefix}{max_num + 1}.{ext}"


def _find_existing_media_by_hash(
    media_dir: Path,
    existing: list[str],
    target_hash: str,
) -> str | None:
    """Find an existing media file with matching hash.

    Args:
        media_dir: Path to ppt/media directory.
        existing: List of existing media filenames.
        target_hash: SHA-256 hash to match against.

    Returns:
        Filename if found, None otherwise.
    """
    for name in existing:
        file_path = media_dir / name

        if (
            file_path.exists()
            and hashlib.sha256(file_path.read_bytes()).hexdigest() == target_hash
        ):
            return name

    return None


def _create_pic_element(
    spid: int,
    name: str,
    media_rid: str,
    audio_rid: str,
    image_rid: str,
    x: int = DEFAULT_ICON_X,
    y: int = DEFAULT_ICON_Y,
    cx: int = DEFAULT_ICON_CX,
    cy: int = DEFAULT_ICON_CY,
) -> ET.Element:
    """Create a p:pic element for the audio icon.

    Args:
        spid: Shape ID for the picture element.
        name: Name attribute for the picture.
        media_rid: Relationship ID for the media file.
        audio_rid: Relationship ID for the audio file.
        image_rid: Relationship ID for the icon image.
        x: X coordinate for icon position (EMU).
        y: Y coordinate for icon position (EMU).
        cx: Width of icon (EMU).
        cy: Height of icon (EMU).

    Returns:
        The created p:pic ElementTree element.
    """
    ET.register_namespace("p", NAMESPACE_P)
    ET.register_namespace("a", NAMESPACE_A)
    ET.register_namespace("r", NAMESPACE_R)
    ET.register_namespace("p14", NAMESPACE_P14)
    ET.register_namespace("a16", NAMESPACE_A16)

    pic = ET.Element(f"{{{NAMESPACE_P}}}pic")

    nv_pic_pr = ET.SubElement(pic, f"{{{NAMESPACE_P}}}nvPicPr")
    c_nv_pr = ET.SubElement(
        nv_pic_pr,
        f"{{{NAMESPACE_P}}}cNvPr",
        id=str(spid),
        name=name,
    )
    ET.SubElement(
        c_nv_pr,
        f"{{{NAMESPACE_A}}}hlinkClick",
        {f"{{{NAMESPACE_R}}}id": "", "action": "ppaction://media"},
    )
    ext_lst = ET.SubElement(c_nv_pr, f"{{{NAMESPACE_A}}}extLst")
    ext = ET.SubElement(
        ext_lst,
        f"{{{NAMESPACE_A}}}ext",
        uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}",
    )
    ET.SubElement(
        ext,
        f"{{{NAMESPACE_A16}}}creationId",
        id=f"{{{str(uuid.uuid4()).upper()}}}",
    )
    c_nv_pic_pr = ET.SubElement(nv_pic_pr, f"{{{NAMESPACE_P}}}cNvPicPr")
    ET.SubElement(c_nv_pic_pr, f"{{{NAMESPACE_A}}}picLocks", noChangeAspect="1")
    nv_pr = ET.SubElement(nv_pic_pr, f"{{{NAMESPACE_P}}}nvPr")
    ET.SubElement(
        nv_pr,
        f"{{{NAMESPACE_A}}}audioFile",
        {f"{{{NAMESPACE_R}}}link": audio_rid},
    )
    nv_ext_lst = ET.SubElement(nv_pr, f"{{{NAMESPACE_P}}}extLst")
    nv_ext = ET.SubElement(
        nv_ext_lst,
        f"{{{NAMESPACE_P}}}ext",
        uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}",
    )
    ET.SubElement(
        nv_ext,
        f"{{{NAMESPACE_P14}}}media",
        {f"{{{NAMESPACE_R}}}embed": media_rid},
    )

    blip_fill = ET.SubElement(pic, f"{{{NAMESPACE_P}}}blipFill")
    ET.SubElement(
        blip_fill,
        f"{{{NAMESPACE_A}}}blip",
        {f"{{{NAMESPACE_R}}}embed": image_rid},
    )
    stretch = ET.SubElement(blip_fill, f"{{{NAMESPACE_A}}}stretch")
    ET.SubElement(stretch, f"{{{NAMESPACE_A}}}fillRect")

    sp_pr = ET.SubElement(pic, f"{{{NAMESPACE_P}}}spPr")
    xfrm = ET.SubElement(sp_pr, f"{{{NAMESPACE_A}}}xfrm")
    ET.SubElement(xfrm, f"{{{NAMESPACE_A}}}off", x=str(x), y=str(y))
    ET.SubElement(xfrm, f"{{{NAMESPACE_A}}}ext", cx=str(cx), cy=str(cy))
    prst_geom = ET.SubElement(sp_pr, f"{{{NAMESPACE_A}}}prstGeom", prst="rect")
    ET.SubElement(prst_geom, f"{{{NAMESPACE_A}}}avLst")

    return pic


def add_audio_to_slide(
    work_path: Path,
    slide_path: str,
    mp3_path: Path,
) -> None:
    """Insert audio into an extracted slide workspace.

    Args:
        work_path: Extracted PPTX workspace root directory.
        slide_path: OOXML slide path (e.g. ppt/slides/slide1.xml).
        mp3_path: Path to the MP3 audio file to insert.

    Raises:
        FileNotFoundError: If workspace directory or MP3 file does not exist.
        SlideXmlNotFoundError: If the slide XML file does not exist.
    """
    if not work_path.exists() or not work_path.is_dir():
        raise FileNotFoundError(f"Workspace not found: {work_path}")

    if not mp3_path.exists():
        raise FileNotFoundError(f"MP3 file not found: {mp3_path}")

    mp3_data = mp3_path.read_bytes()
    mp3_hash = hashlib.sha256(mp3_data).hexdigest()

    ET.register_namespace("", NAMESPACE_CT)
    ET.register_namespace("p", NAMESPACE_P)
    ET.register_namespace("a", NAMESPACE_A)
    ET.register_namespace("p14", NAMESPACE_P14)
    ET.register_namespace("a16", NAMESPACE_A16)

    slide_file_path = work_path / slide_path

    if not slide_file_path.exists():
        raise SlideXmlNotFoundError(slide_path)

    media_dir = work_path / "ppt/media"
    media_dir.mkdir(parents=True, exist_ok=True)

    existing_mp3s = _find_media_files(media_dir, "media", "mp3")
    existing_pngs = _find_media_files(media_dir, "image", "png")

    mp3_filename = _find_existing_media_by_hash(media_dir, existing_mp3s, mp3_hash)
    icon_filename = _find_existing_media_by_hash(
        media_dir,
        existing_pngs,
        AUDIO_ICON_HASH,
    )

    if mp3_filename is None:
        mp3_filename = _next_media_filename(existing_mp3s, "media", "mp3")
        (media_dir / mp3_filename).write_bytes(mp3_data)

    if icon_filename is None:
        icon_filename = _next_media_filename(existing_pngs, "image", "png")
        (media_dir / icon_filename).write_bytes(AUDIO_ICON_BYTES)

    ct_path = work_path / "[Content_Types].xml"
    ct_root = ET.fromstring(ct_path.read_bytes())
    ensure_content_type_default(ct_root, "mp3", "audio/mpeg")
    ensure_content_type_default(ct_root, "png", "image/png")
    ct_path.write_bytes(ET.tostring(ct_root, encoding="UTF-8", xml_declaration=True))

    rels_path = slide_rels_path(work_path, slide_path)
    rels_path.parent.mkdir(parents=True, exist_ok=True)

    if rels_path.exists():
        rels_root = ET.fromstring(rels_path.read_bytes())
    else:
        rels_root = ET.Element(f"{{{NAMESPACE_RELS}}}Relationships")

    media_target = f"../media/{mp3_filename}"
    icon_target = f"../media/{icon_filename}"

    media_rid = find_relationship_by_type_and_target(
        rels_root,
        REL_TYPE_MEDIA,
        media_target,
    )
    audio_rid = find_relationship_by_type_and_target(
        rels_root,
        REL_TYPE_AUDIO,
        media_target,
    )
    image_rid = find_relationship_by_type_and_target(
        rels_root,
        REL_TYPE_IMAGE,
        icon_target,
    )

    if media_rid is None:
        media_rid = add_relationship(rels_root, REL_TYPE_MEDIA, media_target)

    if audio_rid is None:
        audio_rid = add_relationship(rels_root, REL_TYPE_AUDIO, media_target)

    if image_rid is None:
        image_rid = add_relationship(rels_root, REL_TYPE_IMAGE, icon_target)

    ET.register_namespace("", NAMESPACE_RELS)
    rels_path.write_bytes(
        ET.tostring(rels_root, encoding="UTF-8", xml_declaration=True)
    )

    slide_root = ET.fromstring(slide_file_path.read_bytes())
    spid = get_next_shape_id(slide_root)

    sp_tree = get_or_create_pic_parent(slide_root)
    pic = _create_pic_element(
        spid=spid,
        name=mp3_path.stem,
        media_rid=media_rid,
        audio_rid=audio_rid,
        image_rid=image_rid,
    )
    sp_tree.append(pic)

    cmd_parent = get_or_create_command_parent(slide_root)
    audio_parent = get_or_create_audio_parent(slide_root)

    cmd_base_id = get_next_timing_id(slide_root)
    delay = compute_next_delay(cmd_parent)
    cmd_node = create_command_node(spid, delay, cmd_base_id)
    cmd_parent.insert(0, cmd_node)

    audio_ctn_id = get_next_timing_id(slide_root)
    audio_node = create_audio_node(spid, audio_ctn_id)
    audio_parent.insert(0, audio_node)

    slide_file_path.write_bytes(
        ET.tostring(slide_root, encoding="UTF-8", xml_declaration=True)
    )
