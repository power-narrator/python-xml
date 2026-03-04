"""Central PPTX file model with slide-level operations."""

import tempfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from typing import Self
from zipfile import ZIP_DEFLATED, ZipFile

from .audio_model import Audio
from .audio_read import load_slide_audio
from .audio_write import delete_slide_audio, upsert_slide_audio
from .exceptions import (
    InvalidPptxError,
    RelationshipTargetNotFoundError,
    SlideNotFoundError,
)
from .namespaces import (
    NAMESPACE_DCTERMS,
    NAMESPACE_XSI,
    REL_TYPE_NOTES_SLIDE,
    REL_TYPE_SLIDE,
)
from .notes import extract_notes_text, write_slide_notes
from .paths import rels_path_for_path, resolve_target_path, slide_rels_path
from .rels import get_relationships_target_by_type, read_rels_path


def _update_core_xml_modified(core_content: bytes) -> bytes:
    """Update dcterms:modified timestamp in core.xml.

    Args:
        core_content: Raw bytes of docProps/core.xml.

    Returns:
        Updated XML bytes.
    """
    ET.register_namespace("dcterms", NAMESPACE_DCTERMS)
    ET.register_namespace("xsi", NAMESPACE_XSI)
    root = ET.fromstring(core_content)
    modified = root.find(f"{{{NAMESPACE_DCTERMS}}}modified")

    if modified is None:
        modified = ET.SubElement(root, f"{{{NAMESPACE_DCTERMS}}}modified")

    now = datetime.now(timezone.utc).replace(microsecond=0)
    modified.text = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    modified.set(f"{{{NAMESPACE_XSI}}}type", "dcterms:W3CDTF")

    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def _update_app_xml_notes_count(app_content: bytes, notes_count: int) -> bytes:
    """Update Notes count in docProps/app.xml.

    Args:
        app_content: Raw bytes of docProps/app.xml.
        notes_count: Number of slides with notes slides.

    Returns:
        Updated XML bytes.
    """
    namespace = (
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    )
    root = ET.fromstring(app_content)
    notes = root.find(f"{{{namespace}}}Notes")

    if notes is None:
        notes = ET.SubElement(root, f"{{{namespace}}}Notes")

    notes.text = str(notes_count)

    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def _count_slides_with_notes(work_dir: Path) -> int:
    """Count slides that reference a notesSlide relationship.

    Args:
        work_dir: Extracted PPTX workspace root.

    Returns:
        Number of slides with notes relationships.
    """
    rels_dir = work_dir / "ppt/slides/_rels"

    if not rels_dir.exists():
        return 0

    count = 0

    for rels_path in rels_dir.glob("slide*.xml.rels"):
        rels_root = ET.fromstring(rels_path.read_bytes())

        if get_relationships_target_by_type(rels_root, REL_TYPE_NOTES_SLIDE):
            count += 1

    return count


class Slide:
    """Slide model backed by an extracted PPTX workspace."""

    def __init__(
        self,
        index: int,
        slide_path: str,
        work_dir: Path,
    ):
        """Initialize a Slide.

        Args:
            index: Zero-based slide index.
            slide_path: OOXML path (e.g. ppt/slides/slide1.xml).
            work_dir: Extracted PPTX workspace directory.
        """
        self.index = index
        self.slide_path = slide_path
        self._work_dir = work_dir
        self._notes = self._read_notes()
        self.notes_changed = False
        self.audio: list[Audio] = load_slide_audio(self._work_dir, self.slide_path)
        self._audio_by_id: dict[str, Audio] = {a.audio_id: a for a in self.audio}

    @property
    def notes(self) -> str:
        """Get current in-memory notes text for this slide."""
        return self._notes

    def set_notes(self, text: str) -> None:
        """Update in-memory notes text and mark slide as changed.

        Args:
            text: New plain text notes.
        """
        if text != self._notes:
            self._notes = text
            self.notes_changed = True

    def save_notes(self) -> None:
        """Persist notes to workspace if slide notes were edited."""
        if not self.notes_changed:
            return

        write_slide_notes(self._work_dir, self.slide_path, self._notes)
        self.notes_changed = False

    def _read_notes(self) -> str:
        """Read notes text from extracted workspace files.

        Returns:
            Notes text as plain string.

        Raises:
            RelsNotFoundError: If slide relationships are missing.
            RelationshipTargetNotFoundError: If notes path target cannot be read.
        """
        rels_path = slide_rels_path(self._work_dir, self.slide_path)
        slide_rels = read_rels_path(rels_path)

        if not (
            notes_targets := get_relationships_target_by_type(
                slide_rels, REL_TYPE_NOTES_SLIDE
            )
        ):
            return ""

        notes_target = notes_targets[0]
        notes_xml_path = resolve_target_path(self.slide_path, notes_target)
        notes_path = self._work_dir / notes_xml_path

        if not notes_path.exists():
            raise RelationshipTargetNotFoundError(
                rels_path_for_path(self.slide_path), notes_xml_path
            )

        notes_element = ET.fromstring(notes_path.read_bytes())
        return extract_notes_text(notes_element)

    def _reload_audio(self) -> None:
        """Reload slide audio from workspace files."""
        self.audio = load_slide_audio(self._work_dir, self.slide_path)
        self._audio_by_id = {item.audio_id: item for item in self.audio}

    def add_audio(self, mp3_path: Path) -> None:
        """Upsert audio for this slide immediately.

        Args:
            mp3_path: Path to MP3 file.

        Raises:
            FileNotFoundError: If MP3 file is missing.
            SlideXmlNotFoundError: If the slide XML file is missing.
        """
        upsert_slide_audio(self._work_dir, self.slide_path, mp3_path)
        self._reload_audio()

    def delete_audio(self, audio_id: str) -> None:
        """Delete one slide audio entry immediately.

        Args:
            audio_id: In-memory audio identifier.
        """
        audio = self._audio_by_id.get(audio_id)

        if audio is None:
            return

        delete_slide_audio(self._work_dir, self.slide_path, audio)
        self._reload_audio()


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

        core_path.write_bytes(_update_core_xml_modified(core_path.read_bytes()))

        app_path = self._work_dir / "docProps/app.xml"
        notes_count = _count_slides_with_notes(self._work_dir)

        if not app_path.exists():
            raise InvalidPptxError(str(self._source_path), "Missing docProps/app.xml")

        app_path.write_bytes(
            _update_app_xml_notes_count(app_path.read_bytes(), notes_count)
        )

        with ZipFile(output_path, "w", ZIP_DEFLATED) as zip_file:
            for file_path in self._work_dir.rglob("*"):
                if file_path.is_file():
                    rel_name = file_path.relative_to(self._work_dir).as_posix()
                    zip_file.write(file_path, rel_name)

    def close(self) -> None:
        """Cleanup temporary workspace."""
        self._temp_dir.cleanup()
