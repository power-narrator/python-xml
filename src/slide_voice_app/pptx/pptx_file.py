"""Central PPTX file model with slide-level operations."""

import tempfile
from pathlib import Path
from typing import Self
from zipfile import ZIP_DEFLATED, ZipFile

from .docprops import (
    count_slides_with_notes,
    update_app_xml_notes_count,
    update_core_xml_modified,
)
from .exceptions import (
    InvalidPptxError,
    SlideNotFoundError,
)
from .namespaces import REL_TYPE_SLIDE
from .paths import resolve_target_path
from .rels import get_relationships_target_by_type, read_rels_path
from .slide import Slide


class PptxFile:
    """Central PPTX file class for notes and audio operations."""

    def __init__(self, source_path: Path, temp_dir: tempfile.TemporaryDirectory[str]):
        """Initialize the PptxFile.

        Args:
            source_path: Path to the source .pptx file.
            temp_dir: Temporary workspace owner.
        """
        self._source_path = source_path
        self._temp_dir = temp_dir
        self._work_dir = Path(temp_dir.name) / "unpacked"
        self.slides: list[Slide] = []

    @classmethod
    def open(cls, path: Path) -> Self:
        """Open a PPTX file into a temporary workspace.

        Args:
            path: Path to .pptx file.

        Returns:
            Opened PptxFile instance.

        Raises:
            FileNotFoundError: If file does not exist.
            InvalidPptxError: If file is not a valid PPTX.
            RelsNotFoundError: If presentation relationships are missing.
        """
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        temp_dir = tempfile.TemporaryDirectory()

        try:
            with ZipFile(path, "r") as zip_file:
                if "ppt/presentation.xml" not in zip_file.namelist():
                    raise InvalidPptxError(str(path), "Missing ppt/presentation.xml")

                zip_file.extractall(Path(temp_dir.name) / "unpacked")
        except InvalidPptxError:
            temp_dir.cleanup()
            raise
        except Exception as e:
            temp_dir.cleanup()
            raise InvalidPptxError(str(path), f"Cannot open as ZIP: {e}") from e

        instance = cls(path, temp_dir)
        instance._load_slides()
        return instance

    def __enter__(self) -> Self:
        """Enter context manager scope."""
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        """Exit context manager scope and cleanup workspace."""
        self.close()

    @property
    def slide_count(self) -> int:
        """Get the number of slides."""
        return len(self.slides)

    def _load_slides(self) -> None:
        """Load ordered slide paths and instantiate Slide objects.

        Raises:
            RelsNotFoundError: If presentation relationships are missing.
        """
        rels = read_rels_path(self._work_dir / "ppt/_rels/presentation.xml.rels")
        targets = get_relationships_target_by_type(rels, REL_TYPE_SLIDE)

        indexed_paths: list[tuple[int, str]] = []

        for target in targets:
            slide_path = resolve_target_path("ppt/presentation.xml", target)
            slide_num = int(slide_path.split("slide")[-1].replace(".xml", ""))
            indexed_paths.append((slide_num, slide_path))

        indexed_paths.sort(key=lambda item: item[0])
        self.slides = [
            Slide(
                index=index,
                slide_path=slide_path,
                work_dir=self._work_dir,
            )
            for index, (_, slide_path) in enumerate(indexed_paths)
        ]

    def _get_slide(self, slide_index: int) -> Slide:
        """Get slide object by zero-based index.

        Args:
            slide_index: Zero-based slide index.

        Returns:
            The selected Slide.

        Raises:
            SlideNotFoundError: If index is out of range.
        """
        if slide_index < 0 or slide_index >= len(self.slides):
            raise SlideNotFoundError(slide_index, len(self.slides))

        return self.slides[slide_index]

    def get_all_slide_notes(self) -> list[str]:
        """Get notes for all slides.

        Raises:
            RelsNotFoundError: If a slide relationships file is missing.
            RelationshipTargetNotFoundError: If a notes path cannot be read for
                a slide.
        """
        return [slide.notes for slide in self.slides]

    def set_slide_notes(self, slide_index: int, notes: str) -> None:
        """Update in-memory notes for a single slide.

        Args:
            slide_index: Zero-based slide index.
            notes: Plain text notes.

        Raises:
            SlideNotFoundError: If slide index is out of range.
        """
        slide = self._get_slide(slide_index)
        slide.set_notes(notes)

    def save_notes(self) -> None:
        """Persist edited slide notes back into the workspace files."""
        for slide in self.slides:
            slide.save_notes()

    def save_audio_for_slide(self, slide_index: int, mp3_path: Path) -> None:
        """Apply audio update for a slide immediately.

        Args:
            slide_index: Zero-based slide index.
            mp3_path: Path to MP3 file.

        Raises:
            SlideNotFoundError: If slide index is out of range.
            FileNotFoundError: If MP3 file is missing.
            SlideXmlNotFoundError: If the slide XML file is missing.
        """
        slide = self._get_slide(slide_index)
        slide.add_audio(mp3_path)

    def export_to(self, output_path: Path) -> None:
        """Export current workspace into a .pptx file.

        Always updates docProps/core.xml modified timestamp if core.xml exists.

        Args:
            output_path: Destination .pptx path.
        """
        self.save_notes()

        core_path = self._work_dir / "docProps/core.xml"

        if not core_path.exists():
            raise InvalidPptxError(str(self._source_path), "Missing docProps/core.xml")

        core_path.write_bytes(update_core_xml_modified(core_path.read_bytes()))

        app_path = self._work_dir / "docProps/app.xml"
        notes_count = count_slides_with_notes(self._work_dir)

        if not app_path.exists():
            raise InvalidPptxError(str(self._source_path), "Missing docProps/app.xml")

        app_path.write_bytes(
            update_app_xml_notes_count(app_path.read_bytes(), notes_count)
        )

        with ZipFile(output_path, "w", ZIP_DEFLATED) as zip_file:
            for file_path in self._work_dir.rglob("*"):
                if file_path.is_file():
                    rel_name = file_path.relative_to(self._work_dir).as_posix()
                    zip_file.write(file_path, rel_name)

    def close(self) -> None:
        """Cleanup temporary workspace."""
        self._temp_dir.cleanup()
