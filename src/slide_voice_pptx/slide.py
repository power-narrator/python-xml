"""Slide model backed by an extracted PPTX workspace."""

import xml.etree.ElementTree as ET
from pathlib import Path

from .audio.audio_model import Audio
from .audio.audio_read import load_slide_audio
from .audio.audio_write import delete_slide_audio, upsert_slide_audio
from .exceptions import RelationshipTargetNotFoundError
from .namespaces import REL_TYPE_NOTES_SLIDE
from .notes import extract_notes_text, write_slide_notes
from .paths import rels_path_for_path, resolve_target_path, slide_rels_path
from .rels import find_relationship_target_by_type, read_rels_path


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
            RelationshipTargetNotFoundError: If notes path target cannot be read.
        """
        rels_path = slide_rels_path(self._work_dir, self.slide_path)
        slide_rels = read_rels_path(rels_path)

        if (
            notes_target := find_relationship_target_by_type(
                slide_rels,
                REL_TYPE_NOTES_SLIDE,
            )
        ) is None:
            return ""

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
