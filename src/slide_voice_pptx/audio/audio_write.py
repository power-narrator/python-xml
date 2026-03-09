"""Slide audio write helpers."""

import xml.etree.ElementTree as ET
from pathlib import Path

from ..exceptions import SlideXmlNotFoundError
from ..namespaces import NSMAP, NSMAP_RELS
from ..paths import resolve_target_path, slide_rels_path
from ..xpath import (
    XPATH_P_AUDIO,
    XPATH_P_PAR,
    XPATH_P_PIC,
    XPATH_P_SPTGT_BY_SPID,
    XPATH_PIC_CNVPR,
    XPATH_RELATIONSHIP_BY_ID,
)
from .audio_embed import add_audio_to_slide
from .audio_model import Audio
from .audio_read import load_slide_audio


def _remove_nodes_with_spid_target(
    slide_root: ET.Element,
    parent_map: dict[ET.Element, ET.Element],
    nodes_xpath: str,
    spid: int,
) -> None:
    """Remove nodes whose subtree contains a target shape ID."""
    sp_tgt_xpath = XPATH_P_SPTGT_BY_SPID.format(spid=spid)

    for node in slide_root.findall(nodes_xpath, namespaces=NSMAP):
        if node.find(sp_tgt_xpath, namespaces=NSMAP) is None:
            continue

        parent = parent_map.get(node)

        if parent is not None:
            parent.remove(node)


def upsert_slide_audio(work_dir: Path, slide_path: str, mp3_path: Path) -> None:
    """Update matching named audio or insert a new slide audio entry.

    Args:
        work_dir: Extracted PPTX workspace root.
        slide_path: Slide OOXML path.
        mp3_path: Source MP3 path.

    Raises:
        FileNotFoundError: If source file does not exist.
        SlideXmlNotFoundError: If the slide XML file does not exist.
    """
    if not mp3_path.exists():
        raise FileNotFoundError(f"MP3 file not found: {mp3_path}")

    audio_name = mp3_path.stem
    existing_audio = next(
        (
            audio
            for audio in load_slide_audio(work_dir, slide_path)
            if audio.name == audio_name
        ),
        None,
    )

    if existing_audio is not None:
        media_path = work_dir / resolve_target_path(slide_path, existing_audio.target)
        media_path.parent.mkdir(parents=True, exist_ok=True)
        media_path.write_bytes(mp3_path.read_bytes())
        return

    add_audio_to_slide(work_dir, slide_path, mp3_path)


def delete_slide_audio(work_dir: Path, slide_path: str, audio: Audio) -> None:
    """Delete one audio object from slide XML and relationships.

    Args:
        work_dir: Extracted PPTX workspace root.
        slide_path: Slide OOXML path.
        audio: Audio entry to remove.

    Raises:
        SlideXmlNotFoundError: If the slide XML file does not exist.
    """
    spid = audio.spid
    slide_file = work_dir / slide_path

    if not slide_file.exists():
        raise SlideXmlNotFoundError(slide_path)

    slide_root = ET.fromstring(slide_file.read_bytes())
    parent_map = {child: parent for parent in slide_root.iter() for child in parent}

    for pic in slide_root.findall(XPATH_P_PIC, namespaces=NSMAP):
        c_nv_pr = pic.find(XPATH_PIC_CNVPR, namespaces=NSMAP)

        if c_nv_pr is None or c_nv_pr.get("id") != str(spid):
            continue

        parent = parent_map.get(pic)

        if parent is not None:
            parent.remove(pic)

    _remove_nodes_with_spid_target(slide_root, parent_map, XPATH_P_PAR, spid)
    _remove_nodes_with_spid_target(slide_root, parent_map, XPATH_P_AUDIO, spid)

    slide_file.write_bytes(
        ET.tostring(slide_root, encoding="UTF-8", xml_declaration=True)
    )

    rels_file = slide_rels_path(work_dir, slide_path)

    if not rels_file.exists():
        return

    rels_root = ET.fromstring(rels_file.read_bytes())
    ids_to_remove = {audio.audio_rid, audio.media_rid, audio.image_rid}

    for rid in ids_to_remove:
        for relationship in rels_root.findall(
            XPATH_RELATIONSHIP_BY_ID.format(rid=rid),
            namespaces=NSMAP_RELS,
        ):
            rels_root.remove(relationship)

    rels_file.write_bytes(
        ET.tostring(rels_root, encoding="UTF-8", xml_declaration=True)
    )
