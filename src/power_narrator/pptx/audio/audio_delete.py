"""Slide audio delete helpers."""

import xml.etree.ElementTree as ET
from pathlib import Path

from power_narrator.pptx.rels import get_relationship_id_target_map

from ..content_types import remove_content_type_default_if_unused
from ..exceptions import AudioNotFoundError, SlideXmlNotFoundError
from ..namespaces import NSMAP, NSMAP_RELS
from ..paths import resolve_target_path, slide_rels_path, source_path_for_rels_path
from ..xpath import (
    XPATH_P_PIC,
    XPATH_P_SPTGT_BY_SPID,
    XPATH_PIC_AUDIO_FILE,
    XPATH_PIC_BLIP,
    XPATH_PIC_CNVPR,
    XPATH_PIC_MEDIA,
    XPATH_RELATIONSHIP_BY_ID,
)
from .audio_read import load_slide_audio
from .audio_timing import get_automatic_command_parent, normalize_command_delays

TMROOT_CHILD_TN_LST_XPATH = ".//p:cTn[@nodeType='tmRoot']/p:childTnLst"
MAINSEQ_CHILD_TN_LST_XPATH = "p:cTn[@nodeType='mainSeq']/p:childTnLst"


def _slide_uses_relationship_id(slide_root: ET.Element, rid: str) -> bool:
    """Return whether the slide XML still references a relationship ID.

    Args:
        slide_root: Root element of the slide XML.
        rid: Relationship ID to search for.

    Returns:
        True when the relationship ID is still referenced in the slide XML.
    """
    return (
        slide_root.find(
            f"{XPATH_P_PIC}/{XPATH_PIC_AUDIO_FILE}[@r:link='{rid}']",
            namespaces=NSMAP,
        )
        is not None
        or slide_root.find(
            f"{XPATH_P_PIC}/{XPATH_PIC_MEDIA}[@r:embed='{rid}']",
            namespaces=NSMAP,
        )
        is not None
        or slide_root.find(
            f"{XPATH_P_PIC}/{XPATH_PIC_BLIP}[@r:embed='{rid}']",
            namespaces=NSMAP,
        )
        is not None
    )


def _remove_empty_timing(slide_root: ET.Element) -> None:
    """Prune empty timing wrappers and remove timing when nothing remains.

    Args:
        slide_root: Root element of the slide XML.
    """
    timing = slide_root.find("p:timing", namespaces=NSMAP)

    if timing is None:
        return

    tmroot_child_tn_lst = slide_root.find(TMROOT_CHILD_TN_LST_XPATH, namespaces=NSMAP)

    if tmroot_child_tn_lst is None or len(tmroot_child_tn_lst) == 0:
        slide_root.remove(timing)


def _remove_main_sequence_nodes_with_spid_target(
    slide_root: ET.Element,
    spid: int,
) -> None:
    """Remove main sequence timing branches targeting a specific shape ID.

    Args:
        slide_root: Root element of the slide XML.
        spid: Shape ID to remove from the main sequence.
    """
    sp_tgt_xpath = XPATH_P_SPTGT_BY_SPID.format(spid=spid)
    seq_parent = slide_root.find(TMROOT_CHILD_TN_LST_XPATH, namespaces=NSMAP)

    if seq_parent is None:
        return

    for seq in list(seq_parent.findall("p:seq", namespaces=NSMAP)):
        child_tn_lst = seq.find(MAINSEQ_CHILD_TN_LST_XPATH, namespaces=NSMAP)

        if child_tn_lst is None:
            continue

        for par in list(child_tn_lst.findall("p:par", namespaces=NSMAP)):
            if par.find(sp_tgt_xpath, namespaces=NSMAP) is None:
                continue

            nested_child_tn_lst = par.find("p:cTn/p:childTnLst", namespaces=NSMAP)

            if nested_child_tn_lst is None:
                child_tn_lst.remove(par)
                continue

            removed_nested_par = False

            for nested_par in list(
                nested_child_tn_lst.findall("p:par", namespaces=NSMAP)
            ):
                if nested_par.find(sp_tgt_xpath, namespaces=NSMAP) is None:
                    continue

                nested_child_tn_lst.remove(nested_par)
                removed_nested_par = True

            if (
                removed_nested_par
                and nested_child_tn_lst.find("p:par", namespaces=NSMAP) is not None
            ):
                continue

            child_tn_lst.remove(par)

        if child_tn_lst.find("p:par", namespaces=NSMAP) is not None:
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
    sp_tgt_xpath = XPATH_P_SPTGT_BY_SPID.format(spid=spid)
    seq_parent = slide_root.find(TMROOT_CHILD_TN_LST_XPATH, namespaces=NSMAP)

    if seq_parent is None:
        return

    for seq in list(seq_parent.findall("p:seq", namespaces=NSMAP)):
        c_tn = seq.find("p:cTn[@nodeType='interactiveSeq']", namespaces=NSMAP)

        if c_tn is None or c_tn.find(sp_tgt_xpath, namespaces=NSMAP) is None:
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
    sp_tgt_xpath = XPATH_P_SPTGT_BY_SPID.format(spid=spid)
    audio_parent = slide_root.find(TMROOT_CHILD_TN_LST_XPATH, namespaces=NSMAP)

    if audio_parent is None:
        return

    for audio in list(audio_parent.findall(".//p:audio", namespaces=NSMAP)):
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

    _remove_main_sequence_nodes_with_spid_target(slide_root, spid)
    _remove_interactive_sequences_with_spid_target(slide_root, spid)
    _remove_audio_nodes_with_spid_target(slide_root, spid)
    command_parent = get_automatic_command_parent(slide_root)

    if command_parent is not None:
        normalize_command_delays(command_parent)

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
