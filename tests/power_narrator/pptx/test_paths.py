from power_narrator.pptx.paths import (
    relative_target_path,
    rels_path_for_path,
    resolve_target_path,
    slide_rels_path,
)


def test_resolve_target_path_relative_path() -> None:
    assert (
        resolve_target_path("ppt/slides/slide1.xml", "../notesSlides/notesSlide1.xml")
        == "ppt/notesSlides/notesSlide1.xml"
    )


def test_rels_path_for_path_builds_expected_path() -> None:
    assert (
        rels_path_for_path("ppt/slides/slide1.xml")
        == "ppt/slides/_rels/slide1.xml.rels"
    )


def test_slide_rels_path_uses_workspace_root(tmp_path) -> None:
    assert slide_rels_path(tmp_path, "ppt/slides/slide1.xml") == (
        tmp_path / "ppt/slides/_rels/slide1.xml.rels"
    )


def test_relative_target_path_between_slide_and_notes() -> None:
    assert (
        relative_target_path("ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml")
        == "../notesSlides/notesSlide1.xml"
    )
