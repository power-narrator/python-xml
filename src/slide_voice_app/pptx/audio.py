"""Slide audio models and helpers."""

import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

from .audio_insert import add_audio_to_slide
from .namespaces import NAMESPACE_R, NSMAP, NSMAP_RELS
from .paths import resolve_target_path, slide_rels_path


@dataclass(slots=True)
class Audio:
    """Represents one audio entry attached to a slide."""

    audio_id: str
    name: str
    audio_rid: str
    media_rid: str
    image_rid: str
    target: str
    spid: int
    from_workspace: bool = True


def load_slide_audio(work_dir: Path, slide_path: str) -> list[Audio]:
    """Load audio entries from a slide and its relationship file.

    Args:
        work_dir: Extracted PPTX workspace root.
        slide_path: Slide OOXML path.

    Returns:
        List of discovered audio entries.
    """
    slide_file = work_dir / slide_path

    if not slide_file.exists():
        return []

    slide_root = ET.fromstring(slide_file.read_bytes())
    audio_entries: list[Audio] = []
    needed_rids: set[str] = set()
    entry_parts: list[tuple[str, str, int, str, str, str]] = []

    for pic in slide_root.findall(
        ".//p:pic[p:nvPicPr/p:nvPr/a:audioFile]",
        namespaces=NSMAP,
    ):
        c_nv_pr = pic.find("p:nvPicPr/p:cNvPr", namespaces=NSMAP)

        if c_nv_pr is None:
            continue

        spid_value = c_nv_pr.get("id")
        spid = int(spid_value) if spid_value and spid_value.isdigit() else None

        name = c_nv_pr.get("name", "")

        audio_file = pic.find("p:nvPicPr/p:nvPr/a:audioFile", namespaces=NSMAP)
        media_el = pic.find(
            "p:nvPicPr/p:nvPr/p:extLst/p:ext/p14:media",
            namespaces=NSMAP,
        )
        blip = pic.find("p:blipFill/a:blip", namespaces=NSMAP)

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

        for rel in rels_root.findall("r:Relationship", namespaces=NSMAP_RELS):
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


def upsert_slide_audio(work_dir: Path, slide_path: str, mp3_path: Path) -> None:
    """Update matching named audio or insert a new slide audio entry.

    Args:
        work_dir: Extracted PPTX workspace root.
        slide_path: Slide OOXML path.
        mp3_path: Source MP3 path.

    Raises:
        FileNotFoundError: If source file does not exist.
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
    """
    if audio.spid < 0:
        return

    spid = audio.spid

    slide_file = work_dir / slide_path

    if not slide_file.exists():
        return

    slide_root = ET.fromstring(slide_file.read_bytes())
    parent_map = {child: parent for parent in slide_root.iter() for child in parent}

    for pic in slide_root.findall(
        f".//p:pic[p:nvPicPr/p:cNvPr[@id='{spid}']]",
        namespaces=NSMAP,
    ):
        if parent := parent_map.get(pic):
            parent.remove(pic)

    for par in slide_root.findall(
        f".//p:par[.//p:spTgt[@spid='{spid}']]",
        namespaces=NSMAP,
    ):
        if parent := parent_map.get(par):
            parent.remove(par)

    for audio_node in slide_root.findall(
        f".//p:audio[.//p:spTgt[@spid='{spid}']]",
        namespaces=NSMAP,
    ):
        if parent := parent_map.get(audio_node):
            parent.remove(audio_node)

    slide_file.write_bytes(
        ET.tostring(slide_root, encoding="UTF-8", xml_declaration=True)
    )

    rels_file = slide_rels_path(work_dir, slide_path)

    if not rels_file.exists():
        return

    rels_root = ET.fromstring(rels_file.read_bytes())
    ids_to_remove = {audio.audio_rid, audio.media_rid, audio.image_rid}

    if not ids_to_remove:
        return

    for rid in ids_to_remove:
        for relationship in rels_root.findall(
            f"r:Relationship[@Id='{rid}']",
            namespaces=NSMAP_RELS,
        ):
            rels_root.remove(relationship)

    rels_file.write_bytes(
        ET.tostring(rels_root, encoding="UTF-8", xml_declaration=True)
    )
