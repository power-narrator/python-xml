import xml.etree.ElementTree as ET
from pathlib import Path
from typing import cast
from zipfile import ZIP_DEFLATED, ZipFile

from slide_voice_pptx.namespaces import (
    NSMAP,
    NSMAP_CT,
    NSMAP_RELS,
    REL_TYPE_AUDIO,
    REL_TYPE_IMAGE,
    REL_TYPE_MEDIA,
    REL_TYPE_NOTES_SLIDE,
)
from slide_voice_pptx.pptx_file import PptxFile
from slide_voice_pptx.xpath import (
    XPATH_CT_DEFAULT_BY_EXTENSION,
    XPATH_CT_OVERRIDE_BY_PATH_NAME,
    XPATH_P_PIC,
    XPATH_P_TIMING,
    XPATH_PIC_AUDIO_FILE,
    XPATH_PIC_BLIP,
    XPATH_PIC_CNVPR,
    XPATH_PIC_MEDIA,
    XPATH_RELATIONSHIP_WITH_ID,
)


SAMPLES_DIR = Path(__file__).resolve().parents[1] / "data" / "pptx_samples"


def _fixture_pptx_path(tmp_path: Path, sample_name: str) -> Path:
    source_dir = SAMPLES_DIR / sample_name
    pptx_path = tmp_path / f"{sample_name}.pptx"

    with ZipFile(pptx_path, "w", ZIP_DEFLATED) as zip_file:
        for file_path in source_dir.rglob("*"):
            if file_path.is_file():
                zip_file.write(file_path, file_path.relative_to(source_dir).as_posix())

    return pptx_path


def _write_mp3(tmp_path: Path, name: str, payload: bytes) -> Path:
    mp3_path = tmp_path / name
    mp3_path.write_bytes(payload)
    return mp3_path


def _read_zip_xml(pptx_path: Path, member: str) -> ET.Element:
    with ZipFile(pptx_path) as zip_file:
        return ET.fromstring(zip_file.read(member))


def _zip_names(pptx_path: Path) -> set[str]:
    with ZipFile(pptx_path) as zip_file:
        return set(zip_file.namelist())


def _media_members(pptx_path: Path) -> set[str]:
    return {name for name in _zip_names(pptx_path) if name.startswith("ppt/media/")}


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

        entries.append(
            {
                "name": c_nv_pr.get("name", ""),
                "spid": spid,
                "interactive": slide_root.find(
                    f".//p:cTn[@nodeType='interactiveSeq']//p:spTgt[@spid='{spid}']",
                    namespaces=NSMAP,
                )
                is not None,
                "mainseq": slide_root.find(
                    f".//p:cTn[@nodeType='mainSeq']//p:spTgt[@spid='{spid}']",
                    namespaces=NSMAP,
                )
                is not None,
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
    mp3_path = _write_mp3(tmp_path, "intro.mp3", b"intro-audio")

    with PptxFile.open(input_path) as pptx:
        pptx.save_audio_for_slide(0, mp3_path)
        assert _audio_names(pptx.get_slides()) == ["intro"]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    content_types_root = _read_zip_xml(output_path, "[Content_Types].xml")
    audio_entries = _audio_entries(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert len(audio_entries) == 1
    assert audio_entries[0]["name"] == "intro"
    assert audio_entries[0]["mainseq"] is True
    assert audio_entries[0]["interactive"] is False
    assert slide_root.find(XPATH_P_TIMING, namespaces=NSMAP) is not None
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
    first_mp3 = _write_mp3(tmp_path, "intro.mp3", b"intro-audio")
    second_mp3 = _write_mp3(tmp_path, "outro.mp3", b"outro-audio")

    with PptxFile.open(input_path) as pptx:
        pptx.save_audio_for_slide(0, first_mp3)
        pptx.save_audio_for_slide(0, second_mp3)
        assert _audio_names(pptx.get_slides()) == ["intro", "outro"]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    audio_entries = _audio_entries(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert [entry["name"] for entry in audio_entries] == ["intro", "outro"]
    assert all(entry["mainseq"] is True for entry in audio_entries)
    assert all(entry["interactive"] is False for entry in audio_entries)
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
        pptx.delete_audio_for_slide(0, "file_example_MP3_700KB")
        assert _audio_names(pptx.get_slides()) == [
            "slide-voice-app",
            "file_example_MP3_2MG",
            "file_example_MP3_1MG",
        ]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    audio_entries = _audio_entries(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert [entry["name"] for entry in audio_entries] == [
        "slide-voice-app",
        "file_example_MP3_2MG",
        "file_example_MP3_1MG",
    ]
    assert [
        entry["name"] for entry in audio_entries if entry["interactive"] is True
    ] == ["file_example_MP3_1MG"]
    assert [entry["name"] for entry in audio_entries if entry["mainseq"] is True] == [
        "slide-voice-app",
        "file_example_MP3_2MG",
    ]
    assert set(targets_by_type[REL_TYPE_AUDIO]) == {
        "../media/media1.mp3",
        "../media/media2.mp3",
        "../media/media3.mp3",
    }
    assert set(targets_by_type[REL_TYPE_MEDIA]) == {
        "../media/media1.mp3",
        "../media/media2.mp3",
        "../media/media3.mp3",
    }
    assert _media_members(output_path) == {
        "ppt/media/image1.png",
        "ppt/media/media1.mp3",
        "ppt/media/media2.mp3",
        "ppt/media/media3.mp3",
    }


def test_delete_automatic_from_two_manual_and_auto_audio_keeps_other_entries(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "2-manual-and-auto-audio")
    output_path = tmp_path / "auto-removed.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, "slide-voice-app")
        assert _audio_names(pptx.get_slides()) == [
            "file_example_MP3_700KB",
            "file_example_MP3_2MG",
            "file_example_MP3_1MG",
        ]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    audio_entries = _audio_entries(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert [entry["name"] for entry in audio_entries] == [
        "file_example_MP3_700KB",
        "file_example_MP3_2MG",
        "file_example_MP3_1MG",
    ]
    assert [
        entry["name"] for entry in audio_entries if entry["interactive"] is True
    ] == ["file_example_MP3_700KB", "file_example_MP3_1MG"]
    assert [entry["name"] for entry in audio_entries if entry["mainseq"] is True] == [
        "file_example_MP3_2MG"
    ]
    assert set(targets_by_type[REL_TYPE_AUDIO]) == {
        "../media/media1.mp3",
        "../media/media2.mp3",
        "../media/media3.mp3",
    }
    assert set(targets_by_type[REL_TYPE_MEDIA]) == {
        "../media/media1.mp3",
        "../media/media2.mp3",
        "../media/media3.mp3",
    }
    assert _media_members(output_path) == {
        "ppt/media/image1.png",
        "ppt/media/media1.mp3",
        "ppt/media/media2.mp3",
        "ppt/media/media3.mp3",
    }


def test_delete_manual_from_one_manual_and_auto_audio_keeps_automatic(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-manual-and-auto-audio")
    output_path = tmp_path / "manual-removed-from-pair.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, "file_example_MP3_700KB")
        assert _audio_names(pptx.get_slides()) == ["slide-voice-app"]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    audio_entries = _audio_entries(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert [entry["name"] for entry in audio_entries] == ["slide-voice-app"]
    assert audio_entries[0]["interactive"] is False
    assert audio_entries[0]["mainseq"] is True
    assert targets_by_type[REL_TYPE_AUDIO] == ["../media/media1.mp3"]
    assert targets_by_type[REL_TYPE_MEDIA] == ["../media/media1.mp3"]
    assert _media_members(output_path) == {
        "ppt/media/image1.png",
        "ppt/media/media1.mp3",
    }


def test_delete_automatic_from_one_manual_and_auto_audio_keeps_manual(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-manual-and-auto-audio")
    output_path = tmp_path / "auto-removed-from-pair.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, "slide-voice-app")
        assert _audio_names(pptx.get_slides()) == ["file_example_MP3_700KB"]
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    audio_entries = _audio_entries(slide_root)
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert [entry["name"] for entry in audio_entries] == ["file_example_MP3_700KB"]
    assert audio_entries[0]["interactive"] is True
    assert audio_entries[0]["mainseq"] is False
    assert targets_by_type[REL_TYPE_AUDIO] == ["../media/media1.mp3"]
    assert targets_by_type[REL_TYPE_MEDIA] == ["../media/media1.mp3"]
    assert _media_members(output_path) == {
        "ppt/media/image1.png",
        "ppt/media/media1.mp3",
    }


def test_delete_last_automatic_audio_removes_all_audio_artifacts(
    tmp_path: Path,
) -> None:
    input_path = _fixture_pptx_path(tmp_path, "1-auto-audio")
    output_path = tmp_path / "all-auto-removed.pptx"

    with PptxFile.open(input_path) as pptx:
        pptx.delete_audio_for_slide(0, "slide-voice-app")
        assert _audio_names(pptx.get_slides()) == []
        pptx.export_to(output_path)

    slide_root = _read_zip_xml(output_path, "ppt/slides/slide1.xml")
    slide_rels_root = _read_zip_xml(output_path, "ppt/slides/_rels/slide1.xml.rels")
    content_types_root = _read_zip_xml(output_path, "[Content_Types].xml")
    targets_by_type = _relationship_targets_by_type(slide_rels_root)

    assert _audio_entries(slide_root) == []
    assert slide_root.find(XPATH_P_TIMING, namespaces=NSMAP) is None
    assert targets_by_type.get(REL_TYPE_AUDIO, []) == []
    assert targets_by_type.get(REL_TYPE_MEDIA, []) == []
    assert targets_by_type.get(REL_TYPE_IMAGE, []) == []
    assert not _has_content_type_default(content_types_root, "mp3")
    assert not _has_content_type_default(content_types_root, "png")
    assert _media_members(output_path) == set()
