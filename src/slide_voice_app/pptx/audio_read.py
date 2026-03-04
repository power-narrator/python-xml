"""Slide audio read helpers."""

import xml.etree.ElementTree as ET
from pathlib import Path

from .audio_model import Audio
from .exceptions import SlideXmlNotFoundError
from .namespaces import NAMESPACE_R, NSMAP, NSMAP_RELS
from .paths import slide_rels_path
from .xpath import (
    XPATH_P_PIC,
    XPATH_PIC_AUDIO_FILE,
    XPATH_PIC_BLIP,
    XPATH_PIC_CNVPR,
    XPATH_PIC_MEDIA,
    XPATH_RELATIONSHIP_WITH_ID,
)


def load_slide_audio(work_dir: Path, slide_path: str) -> list[Audio]:
    """Load audio entries from a slide and its relationship file.

    Args:
        work_dir: Extracted PPTX workspace root.
        slide_path: Slide OOXML path.

    Returns:
        List of discovered audio entries.

    Raises:
        SlideXmlNotFoundError: If the slide XML file does not exist.
    """
    slide_file = work_dir / slide_path

    if not slide_file.exists():
        raise SlideXmlNotFoundError(slide_path)

    slide_root = ET.fromstring(slide_file.read_bytes())
    audio_entries: list[Audio] = []
    needed_rids: set[str] = set()
    entry_parts: list[tuple[str, str, int, str, str, str]] = []

    for pic in slide_root.findall(XPATH_P_PIC, namespaces=NSMAP):
        audio_file = pic.find(XPATH_PIC_AUDIO_FILE, namespaces=NSMAP)

        if audio_file is None:
            continue

        c_nv_pr = pic.find(XPATH_PIC_CNVPR, namespaces=NSMAP)

        if c_nv_pr is None:
            continue

        spid_value = c_nv_pr.get("id")
        spid = int(spid_value) if spid_value and spid_value.isdigit() else None

        name = c_nv_pr.get("name", "")

        media_el = pic.find(
            XPATH_PIC_MEDIA,
            namespaces=NSMAP,
        )
        blip = pic.find(XPATH_PIC_BLIP, namespaces=NSMAP)

        audio_rid = (
            audio_file.get(f"{{{NAMESPACE_R}}}link") if audio_file is not None else None
        )
        media_rid = (
            media_el.get(f"{{{NAMESPACE_R}}}embed") if media_el is not None else None
        )
        image_rid = blip.get(f"{{{NAMESPACE_R}}}embed") if blip is not None else None
        audio_id = f"spid:{spid}"

        if audio_rid is None or media_rid is None or image_rid is None or spid is None:
            continue

        needed_rids.add(audio_rid)
        needed_rids.add(media_rid)

        entry_parts.append(
            (
                audio_id,
                name,
                spid,
                audio_rid,
                media_rid,
                image_rid,
            )
        )

    rels_targets: dict[str, str] = {}
    rels_file = slide_rels_path(work_dir, slide_path)

    if needed_rids and rels_file.exists():
        rels_root = ET.fromstring(rels_file.read_bytes())

        for rel in rels_root.findall(XPATH_RELATIONSHIP_WITH_ID, namespaces=NSMAP_RELS):
            rid = rel.get("Id")
            target = rel.get("Target")

            if rid and target and rid in needed_rids:
                rels_targets[rid] = target

    for audio_id, name, spid, audio_rid, media_rid, image_rid in entry_parts:
        target = rels_targets.get(audio_rid)

        if target is None:
            target = rels_targets.get(media_rid)

        if target is None:
            continue

        audio_entries.append(
            Audio(
                audio_id=audio_id,
                name=name,
                spid=spid,
                audio_rid=audio_rid,
                media_rid=media_rid,
                image_rid=image_rid,
                target=target,
                from_workspace=True,
            )
        )

    return audio_entries
