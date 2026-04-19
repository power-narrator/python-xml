import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from typing import cast
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from mutagen.mp3 import MP3

from power_narrator.pptx.namespaces import (
    NSMAP,
    NSMAP_CT,
    NSMAP_RELS,
    REL_TYPE_AUDIO,
    REL_TYPE_IMAGE,
    REL_TYPE_MEDIA,
    REL_TYPE_NOTES_SLIDE,
)
from power_narrator.pptx.paths import resolve_target_path
from power_narrator.pptx.pptx_file import PptxFile
from power_narrator.pptx.xpath import (
    XPATH_CT_DEFAULT_BY_EXTENSION,
    XPATH_CT_OVERRIDE_BY_PATH_NAME,
    XPATH_P_PIC,
    XPATH_PIC_AUDIO_FILE,
    XPATH_PIC_BLIP,
    XPATH_PIC_CNVPR,
    XPATH_PIC_MEDIA,
    XPATH_RELATIONSHIP_WITH_ID,
)

ROOT_DIR = Path(__file__).resolve().parents[3]
SAMPLE_DIRS = [ROOT_DIR, ROOT_DIR / "tests" / "data" / "pptx_samples"]


def _sample_dir(sample_name: str) -> Path:
    for samples_dir in SAMPLE_DIRS:
        candidate = samples_dir / sample_name

        if candidate.is_dir():
            return candidate

    raise FileNotFoundError(f"Sample not found: {sample_name}")


def _fixture_pptx_path(tmp_path: Path, sample_name: str) -> Path:
    source_dir = _sample_dir(sample_name)
    pptx_path = tmp_path / f"{sample_name}.pptx"

    with ZipFile(pptx_path, "w", ZIP_DEFLATED) as zip_file:
        for file_path in source_dir.rglob("*"):
            if file_path.is_file():
                zip_file.write(file_path, file_path.relative_to(source_dir).as_posix())

    return pptx_path


def _write_mp3_from_sample(
    tmp_path: Path,
    name: str,
    sample_name: str,
    media_name: str,
) -> Path:
    mp3_path = tmp_path / name
    mp3_path.write_bytes(
        (_sample_dir(sample_name) / "ppt" / "media" / media_name).read_bytes()
    )
    return mp3_path


def _mp3_duration_ms(mp3_path: Path) -> int:
    info = MP3(mp3_path).info

    if info is None or not hasattr(info, "length"):
        raise ValueError(f"Unable to read MP3 duration: {mp3_path}")

    return round(info.length * 1000)


def _read_zip_xml(pptx_path: Path, member: str) -> ET.Element:
    with ZipFile(pptx_path) as zip_file:
        return ET.fromstring(zip_file.read(member))


def _zip_names(pptx_path: Path) -> set[str]:
    with ZipFile(pptx_path) as zip_file:
        return set(zip_file.namelist())


def _media_members(pptx_path: Path) -> set[str]:
    return {name for name in _zip_names(pptx_path) if name.startswith("ppt/media/")}


def _audio_mode_for_spid(slide_root: ET.Element, spid: str) -> str:
    if (
        slide_root.find(
            f".//p:cTn[@nodeType='interactiveSeq']//p:spTgt[@spid='{spid}']",
            namespaces=NSMAP,
        )
        is not None
    ):
        return "interactive"

    for par in slide_root.findall(
        ".//p:cTn[@nodeType='mainSeq']/p:childTnLst/p:par",
        namespaces=NSMAP,
    ):
        if par.find(f".//p:spTgt[@spid='{spid}']", namespaces=NSMAP) is None:
            continue

        st_cond_lst = par.find("p:cTn/p:stCondLst", namespaces=NSMAP)

        if (
            st_cond_lst is not None
            and st_cond_lst.find(
                "p:cond[@evt='onBegin'][@delay='0']",
                namespaces=NSMAP,
            )
            is not None
        ):
            return "auto"

        return "click"

    return "unknown"


def _audio_entries(slide_root: ET.Element) -> list[dict[str, object]]:
    entries: list[dict[str, object]] = []

    for pic in slide_root.findall(XPATH_P_PIC, namespaces=NSMAP):
        c_nv_pr = pic.find(XPATH_PIC_CNVPR, namespaces=NSMAP)
        audio_file = pic.find(XPATH_PIC_AUDIO_FILE, namespaces=NSMAP)
        media = pic.find(XPATH_PIC_MEDIA, namespaces=NSMAP)
        blip = pic.find(XPATH_PIC_BLIP, namespaces=NSMAP)

        if c_nv_pr is None or audio_file is None or media is None or blip is None:
            continue

        spid = c_nv_pr.get("id")

        if spid is None:
            continue

        mode = _audio_mode_for_spid(slide_root, spid)

        entries.append(
            {
                "name": c_nv_pr.get("name", ""),
                "spid": spid,
                "mode": mode,
                "interactive": mode == "interactive",
                "mainseq": mode in {"auto", "click"},
                "audio_rid": audio_file.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link"
                ),
                "media_rid": media.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                ),
                "image_rid": blip.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                ),
            }
        )

    return entries


def _relationship_targets_by_type(rels_root: ET.Element) -> dict[str, list[str]]:
    targets: dict[str, list[str]] = {}

    for relationship in rels_root.findall(
        XPATH_RELATIONSHIP_WITH_ID, namespaces=NSMAP_RELS
    ):
        rel_type = relationship.get("Type")
        target = relationship.get("Target")

        if rel_type is None or target is None:
            continue

        targets.setdefault(rel_type, []).append(target)

    return targets


def _relationship_targets_by_id(rels_root: ET.Element) -> dict[str, str]:
    targets: dict[str, str] = {}

    for relationship in rels_root.findall(
        XPATH_RELATIONSHIP_WITH_ID, namespaces=NSMAP_RELS
    ):
        rel_id = relationship.get("Id")
        target = relationship.get("Target")

        if rel_id is None or target is None:
            continue

        targets[rel_id] = target

    return targets


def _has_content_type_default(content_types_root: ET.Element, extension: str) -> bool:
    return (
        content_types_root.find(
            XPATH_CT_DEFAULT_BY_EXTENSION.format(extension=extension),
            namespaces=NSMAP_CT,
        )
        is not None
    )


def _has_content_type_override(content_types_root: ET.Element, part_name: str) -> bool:
    return (
        content_types_root.find(
            XPATH_CT_OVERRIDE_BY_PATH_NAME.format(path_name=part_name),
            namespaces=NSMAP_CT,
        )
        is not None
    )


def _audio_names(slides: list[dict[str, object]]) -> list[str]:
    audio_list = cast(list[dict[str, str]], slides[0]["audio"])
    return [audio["name"] for audio in audio_list]


def _audio_signatures(slide_root: ET.Element) -> list[tuple[str, str]]:
    return sorted(
        (cast(str, entry["name"]), cast(str, entry["mode"]))
        for entry in _audio_entries(slide_root)
    )


def _automatic_command_timings(slide_root: ET.Element) -> list[dict[str, int]]:
    for par in slide_root.findall(
        ".//p:cTn[@nodeType='mainSeq']/p:childTnLst/p:par",
        namespaces=NSMAP,
    ):
        if (
            par.find(
                "p:cTn/p:stCondLst/p:cond[@evt='onBegin'][@delay='0']",
                namespaces=NSMAP,
            )
            is None
        ):
            continue

        command_parent = par.find("p:cTn/p:childTnLst", namespaces=NSMAP)

        if command_parent is None:
            return []

        timings: list[dict[str, int]] = []

        for command_par in command_parent.findall("p:par", namespaces=NSMAP):
            delay_node = command_par.find(
                "p:cTn/p:stCondLst/p:cond[@delay]",
                namespaces=NSMAP,
            )
            duration_node = command_par.find(
                ".//p:cmd/p:cBhvr/p:cTn",
                namespaces=NSMAP,
            )
            sp_tgt = command_par.find(
                ".//p:cmd/p:cBhvr/p:tgtEl/p:spTgt",
                namespaces=NSMAP,
            )

            if delay_node is None or duration_node is None or sp_tgt is None:
                continue

            delay = delay_node.get("delay", "")
            duration = duration_node.get("dur", "")
            spid = sp_tgt.get("spid", "")

            if not delay.isdigit() or not duration.isdigit() or not spid.isdigit():
                continue

            timings.append(
                {
                    "spid": int(spid),
                    "delay": int(delay),
                    "dur": int(duration),
                }
            )

        return timings

    return []


def _relationship_type_counts(rels_root: ET.Element) -> dict[str, int]:
    return {
        rel_type: len(targets)
        for rel_type, targets in _relationship_targets_by_type(rels_root).items()
    }


def _media_extension_counts_from_zip(pptx_path: Path) -> dict[str, int]:
    return dict(Counter(Path(name).suffix for name in _media_members(pptx_path)))


def _media_extension_counts_from_sample(sample_name: str) -> dict[str, int]:
    media_dir = _sample_dir(sample_name) / "ppt" / "media"

    if not media_dir.exists():
        return {}

    return dict(
        Counter(
            file_path.suffix for file_path in media_dir.iterdir() if file_path.is_file()
        )
    )


def _main_sequence_branch_counts(slide_root: ET.Element) -> dict[str, int]:
    counts: Counter[str] = Counter()

    for par in slide_root.findall(
        ".//p:cTn[@nodeType='mainSeq']/p:childTnLst/p:par",
        namespaces=NSMAP,
    ):
        st_cond_lst = par.find("p:cTn/p:stCondLst", namespaces=NSMAP)

        if (
            st_cond_lst is not None
            and st_cond_lst.find(
                "p:cond[@evt='onBegin'][@delay='0']",
                namespaces=NSMAP,
            )
            is not None
        ):
            counts["auto"] += 1
        else:
            counts["click"] += 1

    return {"auto": counts["auto"], "click": counts["click"]}


def _slide_structure_summary(slide_root: ET.Element) -> dict[str, object]:
    return {
        "timing": slide_root.find("p:timing", namespaces=NSMAP) is not None,
        "audio_signatures": _audio_signatures(slide_root),
        "audio_node_count": len(slide_root.findall(".//p:audio", namespaces=NSMAP)),
        "branch_counts": _main_sequence_branch_counts(slide_root),
        "interactive_seq_count": len(
            slide_root.findall(".//p:cTn[@nodeType='interactiveSeq']", namespaces=NSMAP)
        ),
    }


def _zip_member_bytes(pptx_path: Path, member: str) -> bytes:
    with ZipFile(pptx_path) as zip_file:
        return zip_file.read(member)


def _audio_media_member_for_name(pptx_path: Path, audio_name: str) -> str:
    slide_root = _read_zip_xml(pptx_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(pptx_path, "ppt/slides/_rels/slide1.xml.rels")
    targets_by_id = _relationship_targets_by_id(slide_rels_root)
    audio_entry = next(
        entry for entry in _audio_entries(slide_root) if entry["name"] == audio_name
    )
    media_target = targets_by_id[cast(str, audio_entry["media_rid"])]
    return resolve_target_path("ppt/slides/slide1.xml", media_target)


def test_set_slide_notes_creates_notes_parts_and_updates_metadata(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "base")
    output_path = tmp_path / "with-notes.pptx"

    with PptxFile.open(input_path) as pptx:
        assert pptx.get_slides() == [{"notes": "", "audio": []}]

        pptx.set_slide_notes(0, "One line")
        pptx.export_to(output_path)

    with PptxFile.open(output_path) as exported:
        assert exported.get_slides() == [{"notes": "One line", "audio": []}]

    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    content_types_root = _read_zip_xml(output_path, "[Content_Types].xml")
    app_root = _read_zip_xml(output_path, "docProps/app.xml")
    presentation_root = _read_zip_xml(output_path, "ppt/presentation.xml")
    app_namespace = (
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    )

    assert _relationship_targets_by_type(slide_rels_root)[REL_TYPE_NOTES_SLIDE] == [
        "../notesSlides/notesSlide1.xml"
    ]
    assert _has_content_type_override(
        content_types_root, "/ppt/notesSlides/notesSlide1.xml"
    )
    assert _has_content_type_override(
        content_types_root, "/ppt/notesMasters/notesMaster1.xml"
    )
    assert _has_content_type_override(content_types_root, "/ppt/theme/theme2.xml")
    assert app_root.find(f"{{{app_namespace}}}Notes") is not None
    assert app_root.findtext(f"{{{app_namespace}}}Notes") == "1"
    assert presentation_root.find(".//p:notesMasterIdLst", namespaces=NSMAP) is not None
    assert "ppt/notesSlides/notesSlide1.xml" in _zip_names(output_path)
    assert "ppt/notesMasters/notesMaster1.xml" in _zip_names(output_path)
    assert "ppt/theme/theme2.xml" in _zip_names(output_path)


def test_save_audio_for_slide_creates_single_autoplay_audio(tmp_path: Path) -> None:
    input_path = _fixture_pptx_path(tmp_path, "base")
    output_path = tmp_path / "one-audio.pptx"
    mp3_path = _write_mp3_from_sample(
        tmp_path,
        "intro.mp3",
        "1-auto-audio",
        "media1.mp3",
    )

    with PptxFile.open(input_path) as pptx:
        pptx.save_audio_for_slide(0, mp3_path)
        assert _audio_names(pptx.get_slides()) == ["intro"]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    content_types_root = _read_zip_xml(output_path, "[Content_Types].xml")
    audio_entries = _audio_entries(slide_root)
    timings = _automatic_command_timings(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert len(audio_entries) == 1
    assert audio_entries[0]["name"] == "intro"
    assert audio_entries[0]["mainseq"] is True
    assert audio_entries[0]["interactive"] is False
    assert slide_root.find("p:timing", namespaces=NSMAP) is not None
    assert timings == [{"spid": 4, "delay": 0, "dur": _mp3_duration_ms(mp3_path)}]
    assert len(targets_by_type[REL_TYPE_AUDIO]) == 1
    assert len(targets_by_type[REL_TYPE_MEDIA]) == 1
    assert len(targets_by_type[REL_TYPE_IMAGE]) == 1
    assert _has_content_type_default(content_types_root, "mp3")
    assert _has_content_type_default(content_types_root, "png")
    assert _media_members(output_path) == {
        "ppt/media/image1.png",
        "ppt/media/media1.mp3",
    }


def test_save_audio_for_slide_twice_keeps_two_autoplay_entries(tmp_path: Path) -> None:
    input_path = _fixture_pptx_path(tmp_path, "base")
    output_path = tmp_path / "two-audio.pptx"
    first_mp3 = _write_mp3_from_sample(
        tmp_path,
        "intro.mp3",
        "1-auto-audio",
        "media1.mp3",
    )
    second_mp3 = _write_mp3_from_sample(
        tmp_path,
        "outro.mp3",
        "1-manual-and-auto-audio",
        "media2.mp3",
    )

    with PptxFile.open(input_path) as pptx:
        pptx.save_audio_for_slide(0, first_mp3)
        pptx.save_audio_for_slide(0, second_mp3)
        assert _audio_names(pptx.get_slides()) == ["intro", "outro"]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    audio_entries = _audio_entries(slide_root)
    audio_names_by_spid = {
        cast(str, entry["spid"]): cast(str, entry["name"]) for entry in audio_entries
    }
    timings = _automatic_command_timings(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert [entry["name"] for entry in audio_entries] == ["intro", "outro"]
    assert all(entry["mainseq"] is True for entry in audio_entries)
    assert all(entry["interactive"] is False for entry in audio_entries)
    assert [
        (audio_names_by_spid[str(item["spid"])], item["delay"], item["dur"])
        for item in timings
    ] == [
        ("outro", 0, _mp3_duration_ms(second_mp3)),
        ("intro", _mp3_duration_ms(second_mp3), _mp3_duration_ms(first_mp3)),
    ]
    assert len(targets_by_type[REL_TYPE_AUDIO]) == 2
    assert len(targets_by_type[REL_TYPE_MEDIA]) == 2
    assert len(targets_by_type[REL_TYPE_IMAGE]) == 1
    assert _media_members(output_path) == {
        "ppt/media/image1.png",
        "ppt/media/media1.mp3",
        "ppt/media/media2.mp3",
    }


def test_delete_manual_from_two_manual_and_auto_audio_keeps_other_entries(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "2-manual-and-auto-audio")
    output_path = tmp_path / "manual-removed.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, "ppt_audio_2")
        assert _audio_names(pptx.get_slides()) == [
            "ppt_audio_1",
            "ppt_audio_3",
            "ppt_audio_4",
        ]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    audio_entries = _audio_entries(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert [entry["name"] for entry in audio_entries] == [
        "ppt_audio_1",
        "ppt_audio_3",
        "ppt_audio_4",
    ]
    assert [entry["name"] for entry in audio_entries if entry["mode"] == "click"] == [
        "ppt_audio_4"
    ]
    assert [entry["name"] for entry in audio_entries if entry["mode"] == "auto"] == [
        "ppt_audio_1",
        "ppt_audio_3",
    ]
    assert len(targets_by_type[REL_TYPE_AUDIO]) == 3
    assert len(targets_by_type[REL_TYPE_MEDIA]) == 3
    assert len(targets_by_type[REL_TYPE_IMAGE]) == 1
    assert _media_extension_counts_from_zip(output_path) == {".mp3": 3, ".png": 1}


def test_delete_automatic_from_two_manual_and_auto_audio_keeps_other_entries(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "2-manual-and-auto-audio")
    output_path = tmp_path / "auto-removed.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, "ppt_audio_1")
        assert _audio_names(pptx.get_slides()) == [
            "ppt_audio_2",
            "ppt_audio_3",
            "ppt_audio_4",
        ]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    audio_entries = _audio_entries(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert [entry["name"] for entry in audio_entries] == [
        "ppt_audio_2",
        "ppt_audio_3",
        "ppt_audio_4",
    ]
    assert [entry["name"] for entry in audio_entries if entry["mode"] == "click"] == [
        "ppt_audio_2",
        "ppt_audio_4",
    ]
    assert [entry["name"] for entry in audio_entries if entry["mode"] == "auto"] == [
        "ppt_audio_3",
    ]
    assert len(targets_by_type[REL_TYPE_AUDIO]) == 3
    assert len(targets_by_type[REL_TYPE_MEDIA]) == 3
    assert len(targets_by_type[REL_TYPE_IMAGE]) == 1
    assert _media_extension_counts_from_zip(output_path) == {".mp3": 3, ".png": 1}


def test_delete_manual_from_one_manual_and_auto_audio_keeps_automatic(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-manual-and-auto-audio")
    output_path = tmp_path / "manual-removed-from-pair.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, "ppt_audio_2")
        assert _audio_names(pptx.get_slides()) == ["ppt_audio_1"]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert _audio_signatures(slide_root) == [("ppt_audio_1", "auto")]
    assert len(targets_by_type[REL_TYPE_AUDIO]) == 1
    assert len(targets_by_type[REL_TYPE_MEDIA]) == 1
    assert len(targets_by_type[REL_TYPE_IMAGE]) == 1
    assert _media_members(output_path) == {
        "ppt/media/image1.png",
        next(iter(set(targets_by_type[REL_TYPE_MEDIA]))).replace("..", "ppt"),
    }


def test_delete_automatic_from_one_manual_and_auto_audio_keeps_manual(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-manual-and-auto-audio")
    output_path = tmp_path / "auto-removed-from-pair.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, "ppt_audio_1")
        assert _audio_names(pptx.get_slides()) == ["ppt_audio_2"]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert _audio_signatures(slide_root) == [("ppt_audio_2", "click")]
    assert len(targets_by_type[REL_TYPE_AUDIO]) == 1
    assert len(targets_by_type[REL_TYPE_MEDIA]) == 1
    assert len(targets_by_type[REL_TYPE_IMAGE]) == 1
    assert _media_members(output_path) == {
        "ppt/media/image1.png",
        next(iter(set(targets_by_type[REL_TYPE_MEDIA]))).replace("..", "ppt"),
    }


def test_delete_last_automatic_audio_removes_all_audio_artifacts(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-auto-audio")
    output_path = tmp_path / "all-auto-removed.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, "ppt_audio_1")
        assert _audio_names(pptx.get_slides()) == []
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    content_types_root = _read_zip_xml(output_path, "[Content_Types].xml")
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert _audio_entries(slide_root) == []
    assert slide_root.find("p:timing", namespaces=NSMAP) is None
    assert targets_by_type.get(REL_TYPE_AUDIO, []) == []
    assert targets_by_type.get(REL_TYPE_MEDIA, []) == []
    assert targets_by_type.get(REL_TYPE_IMAGE, []) == []
    assert not _has_content_type_default(content_types_root, "mp3")
    assert not _has_content_type_default(content_types_root, "png")
    assert _media_members(output_path) == set()


def test_open_mixed_timing_sample_reads_all_audio_entries(tmp_path: Path) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-auto-click-interactive-audio")
    slide_root = _read_zip_xml(input_path, "ppt/slides/slide1.xml")

    with PptxFile.open(input_path) as pptx:
        assert _audio_names(pptx.get_slides()) == [
            "ppt_audio_1",
            "ppt_audio_3",
            "ppt_audio_5",
        ]

    assert _audio_signatures(slide_root) == [
        ("ppt_audio_1", "auto"),
        ("ppt_audio_3", "click"),
        ("ppt_audio_5", "interactive"),
    ]


def test_open_double_mixed_timing_sample_reads_all_audio_entries(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "2-auto-click-interactive-audio")
    slide_root = _read_zip_xml(input_path, "ppt/slides/slide1.xml")

    with PptxFile.open(input_path) as pptx:
        assert _audio_names(pptx.get_slides()) == [
            "ppt_audio_1",
            "ppt_audio_2",
            "ppt_audio_3",
            "ppt_audio_4",
            "ppt_audio_5",
            "ppt_audio_6",
        ]

    assert Counter(mode for _, mode in _audio_signatures(slide_root)) == Counter(
        {"auto": 2, "click": 2, "interactive": 2}
    )


def test_save_audio_for_slide_adds_default_auto_to_mixed_timing_slide(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-auto-click-interactive-audio")
    output_path = tmp_path / "mixed-plus-auto.pptx"
    mp3_path = _write_mp3_from_sample(
        tmp_path,
        "narration.mp3",
        "1-manual-and-auto-audio",
        "media2.mp3",
    )
    input_slide_rels_root = _read_zip_xml(
        input_path, "ppt/slides/_rels/slide1.xml.rels"
    )
    input_type_counts = _relationship_type_counts(input_slide_rels_root)
    input_media_counts = _media_extension_counts_from_zip(input_path)

    with PptxFile.open(input_path) as pptx:
        pptx.save_audio_for_slide(0, mp3_path)
        assert _audio_names(pptx.get_slides()) == [
            "ppt_audio_1",
            "ppt_audio_3",
            "ppt_audio_5",
            "narration",
        ]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    audio_names_by_spid = {
        cast(str, entry["spid"]): cast(str, entry["name"])
        for entry in _audio_entries(slide_root)
    }

    assert Counter(mode for _, mode in _audio_signatures(slide_root)) == Counter(
        {"auto": 2, "click": 1, "interactive": 1}
    )
    assert [
        (audio_names_by_spid[str(item["spid"])], item["delay"], item["dur"])
        for item in _automatic_command_timings(slide_root)
    ] == [
        ("narration", 0, _mp3_duration_ms(mp3_path)),
        ("ppt_audio_1", _mp3_duration_ms(mp3_path), 1224),
    ]
    assert _relationship_type_counts(slide_rels_root)[REL_TYPE_AUDIO] == 4
    assert _relationship_type_counts(slide_rels_root)[REL_TYPE_MEDIA] == 4
    assert _relationship_type_counts(slide_rels_root)[REL_TYPE_IMAGE] in {
        input_type_counts[REL_TYPE_IMAGE],
        input_type_counts[REL_TYPE_IMAGE] + 1,
    }
    assert _media_extension_counts_from_zip(output_path)[".mp3"] == (
        input_media_counts[".mp3"] + 1
    )


@pytest.mark.parametrize(
    ("audio_name", "expected_mode"),
    [
        ("ppt_audio_1", "auto"),
        ("ppt_audio_3", "click"),
        ("ppt_audio_5", "interactive"),
    ],
)
def test_save_audio_for_slide_updates_existing_audio_without_changing_structure(
    tmp_path: Path,
    audio_name: str,
    expected_mode: str,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-auto-click-interactive-audio")
    output_path = tmp_path / f"updated-{audio_name}.pptx"
    mp3_path = _write_mp3_from_sample(
        tmp_path,
        f"{audio_name}.mp3",
        "1-manual-and-auto-audio",
        "media2.mp3",
    )

    with PptxFile.open(input_path) as pptx:
        pptx.save_audio_for_slide(0, mp3_path)
        assert _audio_names(pptx.get_slides()) == [
            "ppt_audio_1",
            "ppt_audio_3",
            "ppt_audio_5",
        ]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    media_member = _audio_media_member_for_name(output_path, audio_name)

    assert (audio_name, expected_mode) in _audio_signatures(slide_root)
    assert Counter(mode for _, mode in _audio_signatures(slide_root)) == Counter(
        {"auto": 1, "click": 1, "interactive": 1}
    )
    assert _relationship_type_counts(slide_rels_root)[REL_TYPE_AUDIO] == 3
    assert _relationship_type_counts(slide_rels_root)[REL_TYPE_MEDIA] == 3
    assert _relationship_type_counts(slide_rels_root)[REL_TYPE_IMAGE] == 1
    assert _zip_member_bytes(output_path, media_member) == mp3_path.read_bytes()


def test_save_audio_for_slide_updates_existing_auto_duration_and_reflows_following_auto(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "2-auto-click-interactive-audio")
    output_path = tmp_path / "updated-auto-timing.pptx"
    mp3_path = _write_mp3_from_sample(
        tmp_path,
        "ppt_audio_1.mp3",
        "1-manual-and-auto-audio",
        "media2.mp3",
    )

    with PptxFile.open(input_path) as pptx:
        pptx.save_audio_for_slide(0, mp3_path)
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    audio_names_by_spid = {
        cast(str, entry["spid"]): cast(str, entry["name"])
        for entry in _audio_entries(slide_root)
    }

    assert [
        (audio_names_by_spid[str(item["spid"])], item["delay"], item["dur"])
        for item in _automatic_command_timings(slide_root)
    ] == [
        ("ppt_audio_1", 0, _mp3_duration_ms(mp3_path)),
        ("ppt_audio_2", _mp3_duration_ms(mp3_path), 1368),
    ]


@pytest.mark.parametrize(
    ("audio_name", "expected_signatures"),
    [
        (
            "ppt_audio_1",
            [
                ("ppt_audio_2", "auto"),
                ("ppt_audio_3", "click"),
                ("ppt_audio_4", "click"),
                ("ppt_audio_5", "interactive"),
                ("ppt_audio_6", "interactive"),
            ],
        ),
        (
            "ppt_audio_3",
            [
                ("ppt_audio_1", "auto"),
                ("ppt_audio_2", "auto"),
                ("ppt_audio_4", "click"),
                ("ppt_audio_5", "interactive"),
                ("ppt_audio_6", "interactive"),
            ],
        ),
        (
            "ppt_audio_5",
            [
                ("ppt_audio_1", "auto"),
                ("ppt_audio_2", "auto"),
                ("ppt_audio_3", "click"),
                ("ppt_audio_4", "click"),
                ("ppt_audio_6", "interactive"),
            ],
        ),
    ],
)
def test_delete_audio_for_slide_keeps_other_entries_in_double_mixed_sample(
    tmp_path: Path,
    audio_name: str,
    expected_signatures: list[tuple[str, str]],
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "2-auto-click-interactive-audio")
    output_path = tmp_path / f"deleted-{audio_name}.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, audio_name)
        assert _audio_names(pptx.get_slides()) == [
            name for name, _ in expected_signatures
        ]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")

    assert _audio_signatures(slide_root) == expected_signatures
    assert _relationship_type_counts(slide_rels_root)[REL_TYPE_AUDIO] == 5
    assert _relationship_type_counts(slide_rels_root)[REL_TYPE_MEDIA] == 5
    assert _relationship_type_counts(slide_rels_root)[REL_TYPE_IMAGE] == 1
    assert _media_extension_counts_from_zip(output_path) == {".mp3": 5, ".png": 1}


@pytest.mark.parametrize(
    ("audio_name", "expected_sample"),
    [
        ("ppt_audio_1", "no-auto-audio"),
        ("ppt_audio_3", "no-click-audio"),
        ("ppt_audio_5", "no-interactive-audio"),
    ],
)
def test_delete_audio_for_slide_cleans_up_timing_groups_for_single_mixed_sample(
    tmp_path: Path,
    audio_name: str,
    expected_sample: str,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-auto-click-interactive-audio")
    output_path = tmp_path / f"{expected_sample}.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, audio_name)
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    content_types_root = _read_zip_xml(output_path, "[Content_Types].xml")
    expected_slide_root = ET.fromstring(
        (_sample_dir(expected_sample) / "ppt/slides/slide1.xml").read_bytes()
    )
    expected_rels_root = ET.fromstring(
        (_sample_dir(expected_sample) / "ppt/slides/_rels/slide1.xml.rels").read_bytes()
    )
    expected_content_types_root = ET.fromstring(
        (_sample_dir(expected_sample) / "[Content_Types].xml").read_bytes()
    )

    assert _slide_structure_summary(slide_root) == _slide_structure_summary(
        expected_slide_root
    )
    assert _relationship_type_counts(slide_rels_root) == _relationship_type_counts(
        expected_rels_root
    )
    assert _media_extension_counts_from_zip(
        output_path
    ) == _media_extension_counts_from_sample(expected_sample)
    assert _has_content_type_default(
        content_types_root, "mp3"
    ) == _has_content_type_default(expected_content_types_root, "mp3")
    assert _has_content_type_default(
        content_types_root, "png"
    ) == _has_content_type_default(expected_content_types_root, "png")
