"""Slide audio delete helpers."""

import xml.etree.ElementTree as ET
from pathlib import Path

from power_narrator.pptx.rels import get_relationship_id_target_map

from ..content_types import remove_content_type_default_if_unused
from ..exceptions import AudioNotFoundError, SlideXmlNotFoundError
from ..namespaces import NAMESPACE_R, NSMAP, NSMAP_RELS
from ..paths import resolve_target_path, slide_rels_path, source_path_for_rels_path
from ..xpath import (
    XPATH_P_AUDIO,
    XPATH_P_MAINSEQ_CHILD_TNLST,
    XPATH_P_PIC,
    XPATH_P_SEQ,
    XPATH_P_SEQ_CHILD,
    XPATH_P_SEQ_INTERACTIVE_CTN_BY_SPID,
    XPATH_P_SEQ_MAINSEQ_CTN,
    XPATH_P_SPTGT_BY_SPID,
    XPATH_P_TIMING,
    XPATH_P_TMROOT_CHILD_TNLST,
    XPATH_PIC_AUDIO_FILE,
    XPATH_PIC_BLIP,
    XPATH_PIC_CNVPR,
    XPATH_PIC_MEDIA,
    XPATH_RELATIONSHIP_BY_ID,
)
from .audio_read import load_slide_audio


def _slide_uses_relationship_id(slide_root: ET.Element, rid: str) -> bool:
    """Return whether the slide XML still references a relationship ID.

    Args:
        slide_root: Root element of the slide XML.
        rid: Relationship ID to search for.

    Returns:
        True when the relationship ID is still referenced in the slide XML.
    """
    rel_attr = f"{{{NAMESPACE_R}}}link"
    embed_attr = f"{{{NAMESPACE_R}}}embed"

    for pic in slide_root.findall(XPATH_P_PIC, namespaces=NSMAP):
        audio_file = pic.find(XPATH_PIC_AUDIO_FILE, namespaces=NSMAP)
        media = pic.find(XPATH_PIC_MEDIA, namespaces=NSMAP)
        blip = pic.find(XPATH_PIC_BLIP, namespaces=NSMAP)

        if audio_file is not None and audio_file.get(rel_attr) == rid:
            return True

        if media is not None and media.get(embed_attr) == rid:
            return True

        if blip is not None and blip.get(embed_attr) == rid:
            return True

    return False


def _remove_empty_timing(slide_root: ET.Element) -> None:
    """Prune empty timing wrappers and remove timing when nothing remains.

    Args:
        slide_root: Root element of the slide XML.
    """
    timing = slide_root.find(XPATH_P_TIMING, namespaces=NSMAP)

    if timing is None:
        return

    if timing.find(XPATH_P_SEQ, namespaces=NSMAP) is None:
        slide_root.remove(timing)


def _remove_non_interactive_sequences_with_spid_target(
    slide_root: ET.Element,
    spid: int,
) -> None:
    """Remove non-interactive sequence wrappers targeting a specific shape ID.

    Args:
        slide_root: Root element of the slide XML.
        spid: Shape ID to remove from non-interactive sequences.
    """
    sp_tgt_xpath = XPATH_P_SPTGT_BY_SPID.format(spid=spid)
    par_parent = slide_root.find(XPATH_P_MAINSEQ_CHILD_TNLST, namespaces=NSMAP)

    if par_parent is None:
        return

    for par in par_parent.findall("p:par", namespaces=NSMAP):
        if par.find(sp_tgt_xpath, namespaces=NSMAP) is None:
            continue

        par_parent.remove(par)

    if len(par_parent) != 0:
        return

    seq_parent = slide_root.find(XPATH_P_TMROOT_CHILD_TNLST, namespaces=NSMAP)

    if seq_parent is None:
        return

    for seq in seq_parent.findall(XPATH_P_SEQ_CHILD, namespaces=NSMAP):
        c_tn = seq.find(XPATH_P_SEQ_MAINSEQ_CTN, namespaces=NSMAP)

        if c_tn is None:
            continue

        seq_parent.remove(seq)


def _remove_interactive_sequences_with_spid_target(
    slide_root: ET.Element,
    spid: int,
) -> None:
    """Remove interactive sequence wrappers targeting a specific shape ID.

    Args:
        slide_root: Root element of the slide XML.
        spid: Shape ID to remove from interactive sequences.
    """
    interactive_seq_xpath = XPATH_P_SEQ_INTERACTIVE_CTN_BY_SPID.format(spid=spid)
    seq_parent = slide_root.find(XPATH_P_TMROOT_CHILD_TNLST, namespaces=NSMAP)

    if seq_parent is None:
        return

    for seq in seq_parent.findall(XPATH_P_SEQ_CHILD, namespaces=NSMAP):
        c_tn = seq.find(interactive_seq_xpath, namespaces=NSMAP)

        if c_tn is None:
            continue

        seq_parent.remove(seq)


def _remove_audio_nodes_with_spid_target(
    slide_root: ET.Element,
    spid: int,
) -> None:
    """Remove audio nodes targeting a specific shape ID.

    Args:
        slide_root: Root element of the slide XML.
        spid: Shape ID to remove from audio timing nodes.
    """
    sp_tgt_xpath = XPATH_P_SPTGT_BY_SPID.format(spid=spid)
    audio_parent = slide_root.find(XPATH_P_TMROOT_CHILD_TNLST, namespaces=NSMAP)

    if audio_parent is None:
        return

    for audio in audio_parent.findall(XPATH_P_AUDIO, namespaces=NSMAP):
        if audio.find(sp_tgt_xpath, namespaces=NSMAP) is None:
            continue

        audio_parent.remove(audio)


def _slides_use_target(work_dir: Path, target_path: str) -> bool:
    """Return whether any slide relationships still resolve to a target path.

    Args:
        work_dir: Extracted PPTX workspace root.
        target_path: Package-relative media target path to check.

    Returns:
        True when at least one slide relationship still resolves to the target path.
    """
    slides_rels_dir = work_dir / "ppt/slides/_rels"

    if not slides_rels_dir.exists():
        return False

    for rels_path in slides_rels_dir.glob("*.rels"):
        rels_part_path = rels_path.relative_to(work_dir).as_posix()
        source_path = source_path_for_rels_path(rels_part_path)
        rels_root = ET.fromstring(rels_path.read_bytes())

        for target in get_relationship_id_target_map(rels_root).values():
            if resolve_target_path(source_path, target) == target_path:
                return True

    return False


def delete_slide_audio(work_dir: Path, slide_path: str, name: str) -> None:
    """Delete the first matching named audio object from slide XML.

    Args:
        work_dir: Extracted PPTX workspace root.
        slide_path: Slide OOXML path.
        name: Audio name to remove.

    Raises:
        AudioNotFoundError: If no audio entry matches the requested name.
        SlideXmlNotFoundError: If the slide XML file does not exist.
    """
    audio = next(
        (item for item in load_slide_audio(work_dir, slide_path) if item.name == name),
        None,
    )

    if audio is None:
        raise AudioNotFoundError(slide_path, name)

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

    _remove_non_interactive_sequences_with_spid_target(slide_root, spid)
    _remove_interactive_sequences_with_spid_target(slide_root, spid)
    _remove_audio_nodes_with_spid_target(slide_root, spid)
    _remove_empty_timing(slide_root)

    slide_file.write_bytes(
        ET.tostring(slide_root, encoding="UTF-8", xml_declaration=True)
    )

    rels_file = slide_rels_path(work_dir, slide_path)

    if not rels_file.exists():
        return

    rels_root = ET.fromstring(rels_file.read_bytes())
    ids_to_remove = {
        rid
        for rid in (audio.audio_rid, audio.media_rid, audio.image_rid)
        if not _slide_uses_relationship_id(slide_root, rid)
    }
    targets_to_check = get_relationship_id_target_map(
        rels_root, only_ids=ids_to_remove
    ).values()

    for rid in ids_to_remove:
        for relationship in rels_root.findall(
            XPATH_RELATIONSHIP_BY_ID.format(rid=rid),
            namespaces=NSMAP_RELS,
        ):
            rels_root.remove(relationship)

    rels_file.write_bytes(
        ET.tostring(rels_root, encoding="UTF-8", xml_declaration=True)
    )
    media_dir = work_dir / "ppt/media"

    for target in targets_to_check:
        target_path = resolve_target_path(slide_path, target)

        if _slides_use_target(work_dir, target_path):
            continue

        absolute_target_path = work_dir / target_path

        if absolute_target_path.exists():
            absolute_target_path.unlink()

    if media_dir.exists():
        remove_content_type_default_if_unused(work_dir, media_dir, "mp3")
        remove_content_type_default_if_unused(work_dir, media_dir, "png")
