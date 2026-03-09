"""Slide audio timing XML helpers."""

import xml.etree.ElementTree as ET

from ..namespaces import NAMESPACE_P, NSMAP
from ..xml_helper import ensure_child
from ..xpath import (
    XPATH_P_CNVPR_WITH_ID,
    XPATH_P_CTN_WITH_ID,
    XPATH_P_SPTGT_WITH_SPID,
    XPATH_TIMING_CONDS_WITH_DELAY,
)

DEFAULT_VOLUME = 80000


def _get_max_shape_id(slide_root: ET.Element) -> int:
    """Scan slide XML for maximum shape/target id values.

    Args:
        slide_root: Root element of the slide XML.

    Returns:
        Maximum shape ID found, or 0 if none found.
    """
    targets = [
        (XPATH_P_CNVPR_WITH_ID, "id"),
        (XPATH_P_SPTGT_WITH_SPID, "spid"),
    ]
    found_ids = [
        int(elem.get(attr, ""))
        for xpath, attr in targets
        for elem in slide_root.findall(xpath, namespaces=NSMAP)
    ]

    return max(found_ids, default=0)


def _get_max_ctn_id(slide_root: ET.Element) -> int:
    """Scan slide timing nodes for maximum cTn id values.

    Args:
        slide_root: Root element of the slide XML.

    Returns:
        Maximum cTn ID found, or 0 if none found.
    """
    ids = (
        int(elem.get("id", ""))
        for elem in slide_root.findall(XPATH_P_CTN_WITH_ID, namespaces=NSMAP)
        if elem.get("id", "").isdigit()
    )

    return max(ids, default=0)


def _get_common_timing_prefix(slide_root: ET.Element) -> ET.Element:
    """Ensure timing prefix exists and return the root childTnLst.

    Path: p:timing/p:tnLst/p:par/p:cTn/p:childTnLst

    Args:
        slide_root: Root element of the slide XML.

    Returns:
        The childTnLst element for appending audio nodes.
    """
    p = NAMESPACE_P
    timing = ensure_child(slide_root, f"{{{p}}}timing", {})
    tn_lst = ensure_child(timing, f"{{{p}}}tnLst", {})
    par = ensure_child(tn_lst, f"{{{p}}}par", {})
    c_tn_root = ensure_child(
        par,
        f"{{{p}}}cTn",
        {
            "dur": "indefinite",
            "restart": "never",
            "nodeType": "tmRoot",
        },
    )

    if c_tn_root.get("id") is None:
        max_id = _get_max_ctn_id(slide_root)
        c_tn_root.set("id", str(max_id + 1))

    return ensure_child(c_tn_root, f"{{{p}}}childTnLst", {})


def get_or_create_command_parent(slide_root: ET.Element) -> ET.Element:
    """Find or create the childTnLst where command nodes should be appended.

    Path: p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/
          p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst

    Args:
        slide_root: Root element of the slide XML.

    Returns:
        The childTnLst element for appending command nodes.
    """
    p = NAMESPACE_P
    root_child_tn_lst = _get_common_timing_prefix(slide_root)

    seq = ensure_child(
        root_child_tn_lst,
        f"{{{p}}}seq",
        {"concurrent": "1", "nextAc": "seek"},
    )
    c_tn_seq = ensure_child(
        seq,
        f"{{{p}}}cTn",
        {"dur": "indefinite", "nodeType": "mainSeq"},
    )

    if c_tn_seq.get("id") is None:
        max_id = _get_max_ctn_id(slide_root)
        c_tn_seq.set("id", str(max_id + 1))

    child_tn_lst = ensure_child(c_tn_seq, f"{{{p}}}childTnLst", {})
    par = ensure_child(child_tn_lst, f"{{{p}}}par", {})
    c_tn_inner = ensure_child(par, f"{{{p}}}cTn", {"fill": "hold"})

    if c_tn_inner.get("id") is None:
        max_id = _get_max_ctn_id(slide_root)
        c_tn_inner.set("id", str(max_id + 1))

    st_cond_lst = ensure_child(c_tn_inner, f"{{{p}}}stCondLst", {})
    ensure_child(st_cond_lst, f"{{{p}}}cond", {"delay": "indefinite"})
    cond_on_begin = ensure_child(
        st_cond_lst,
        f"{{{p}}}cond",
        {"evt": "onBegin", "delay": "0"},
    )
    ensure_child(cond_on_begin, f"{{{p}}}tn", {"val": c_tn_seq.get("id", "")})

    command_parent = ensure_child(c_tn_inner, f"{{{p}}}childTnLst", {})

    prev_cond_lst = ensure_child(seq, f"{{{p}}}prevCondLst", {})
    cond_prev = ensure_child(
        prev_cond_lst,
        f"{{{p}}}cond",
        {"evt": "onPrev", "delay": "0"},
    )
    tgt_prev = ensure_child(cond_prev, f"{{{p}}}tgtEl", {})
    ensure_child(tgt_prev, f"{{{p}}}sldTgt", {})

    next_cond_lst = ensure_child(seq, f"{{{p}}}nextCondLst", {})
    cond_next = ensure_child(
        next_cond_lst,
        f"{{{p}}}cond",
        {"evt": "onNext", "delay": "0"},
    )
    tgt_next = ensure_child(cond_next, f"{{{p}}}tgtEl", {})
    ensure_child(tgt_next, f"{{{p}}}sldTgt", {})

    return command_parent


def get_or_create_audio_parent(slide_root: ET.Element) -> ET.Element:
    """Find or create the childTnLst where audio nodes should be appended.

    Path: p:timing/p:tnLst/p:par/p:cTn/p:childTnLst

    Args:
        slide_root: Root element of the slide XML.

    Returns:
        The childTnLst element for appending audio nodes.
    """
    return _get_common_timing_prefix(slide_root)


def get_or_create_pic_parent(slide_root: ET.Element) -> ET.Element:
    """Find or create the spTree where pic nodes should be appended.

    Path: p:cSld/p:spTree

    Args:
        slide_root: Root element of the slide XML.

    Returns:
        The spTree element for appending pic nodes.
    """
    p = NAMESPACE_P
    c_sld = ensure_child(slide_root, f"{{{p}}}cSld", {})
    return ensure_child(c_sld, f"{{{p}}}spTree", {})


def compute_next_delay(cmd_parent: ET.Element) -> int:
    """Compute delay for new command based on existing playFrom commands.

    Delay increments by 1 for each existing command.

    Args:
        cmd_parent: The childTnLst element containing command nodes.

    Returns:
        Delay for the next audio command.
    """
    max_delay = -1

    for cond in cmd_parent.findall(XPATH_TIMING_CONDS_WITH_DELAY, namespaces=NSMAP):
        delay_value = cond.get("delay", "")

        if delay_value.isdigit():
            max_delay = max(max_delay, int(delay_value))

    return max_delay + 1


def create_command_node(spid: int, delay: int, base_id: int) -> ET.Element:
    """Create a p:par command node for autoplay.

    Args:
        spid: Shape ID to target with the command.
        delay: The index of command within its parent.
        base_id: Starting ID for timing node IDs.

    Returns:
        The created command node element.
    """
    p = NAMESPACE_P

    par = ET.Element(f"{{{p}}}par")
    c_tn_outer = ET.SubElement(par, f"{{{p}}}cTn", id=str(base_id), fill="hold")

    st_cond_lst = ET.SubElement(c_tn_outer, f"{{{p}}}stCondLst")
    ET.SubElement(st_cond_lst, f"{{{p}}}cond", delay=str(delay))

    child_tn_lst = ET.SubElement(c_tn_outer, f"{{{p}}}childTnLst")
    inner_par = ET.SubElement(child_tn_lst, f"{{{p}}}par")
    c_tn_inner = ET.SubElement(
        inner_par,
        f"{{{p}}}cTn",
        id=str(base_id + 1),
        presetID="1",
        presetClass="mediacall",
        presetSubtype="0",
        fill="hold",
        nodeType="afterEffect",
    )

    st_cond_lst_inner = ET.SubElement(c_tn_inner, f"{{{p}}}stCondLst")
    ET.SubElement(st_cond_lst_inner, f"{{{p}}}cond", delay="0")

    child_tn_lst_inner = ET.SubElement(c_tn_inner, f"{{{p}}}childTnLst")
    cmd = ET.SubElement(
        child_tn_lst_inner,
        f"{{{p}}}cmd",
        type="call",
        cmd="playFrom(0.0)",
    )

    c_bhvr = ET.SubElement(cmd, f"{{{p}}}cBhvr")
    ET.SubElement(
        c_bhvr,
        f"{{{p}}}cTn",
        id=str(base_id + 2),
        dur="1",
        fill="hold",
    )
    tgt_el = ET.SubElement(c_bhvr, f"{{{p}}}tgtEl")
    ET.SubElement(tgt_el, f"{{{p}}}spTgt", spid=str(spid))

    return par


def create_audio_node(
    spid: int,
    timing_id: int,
    volume: int = DEFAULT_VOLUME,
) -> ET.Element:
    """Create a p:audio node for the media.

    Args:
        spid: Shape ID of the audio icon.
        timing_id: ID for the timing node.
        volume: Audio volume level.

    Returns:
        The created audio node element.
    """
    p = NAMESPACE_P

    audio = ET.Element(f"{{{p}}}audio")
    c_media_node = ET.SubElement(
        audio,
        f"{{{p}}}cMediaNode",
        vol=str(volume),
        showWhenStopped="0",
    )

    c_tn = ET.SubElement(
        c_media_node,
        f"{{{p}}}cTn",
        id=str(timing_id),
        fill="hold",
        display="0",
    )

    st_cond_lst = ET.SubElement(c_tn, f"{{{p}}}stCondLst")
    ET.SubElement(st_cond_lst, f"{{{p}}}cond", delay="indefinite")

    end_cond_lst = ET.SubElement(c_tn, f"{{{p}}}endCondLst")
    cond = ET.SubElement(end_cond_lst, f"{{{p}}}cond", evt="onStopAudio", delay="0")
    tgt_el = ET.SubElement(cond, f"{{{p}}}tgtEl")
    ET.SubElement(tgt_el, f"{{{p}}}sldTgt")

    tgt_el_2 = ET.SubElement(c_media_node, f"{{{p}}}tgtEl")
    ET.SubElement(tgt_el_2, f"{{{p}}}spTgt", spid=str(spid))

    return audio


def get_next_shape_id(slide_root: ET.Element) -> int:
    """Get next available shape id from slide content."""
    return _get_max_shape_id(slide_root) + 1


def get_next_timing_id(slide_root: ET.Element) -> int:
    """Get next available timing id from slide timing nodes."""
    return _get_max_ctn_id(slide_root) + 1
