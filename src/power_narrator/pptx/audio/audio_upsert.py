"""Slide audio upsert helpers."""

import xml.etree.ElementTree as ET
from pathlib import Path

from ..exceptions import SlideXmlNotFoundError
from ..paths import resolve_target_path
from .audio_insert import add_audio_to_slide, get_mp3_duration_ms
from .audio_read import load_slide_audio
from .audio_timing import update_automatic_command_duration


def upsert_slide_audio(work_dir: Path, slide_path: str, mp3_path: Path) -> None:
    """Update matching named audio or insert a new slide audio entry.

    Args:
        work_dir: Extracted PPTX workspace root.
        slide_path: Slide OOXML path.
        mp3_path: Source MP3 path.

    Raises:
        FileNotFoundError: If source file does not exist.
        SlideXmlNotFoundError: If the slide XML file does not exist.
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

        slide_file = work_dir / slide_path

        if not slide_file.exists():
            raise SlideXmlNotFoundError(slide_path)

        slide_root = ET.fromstring(slide_file.read_bytes())

        if update_automatic_command_duration(
            slide_root,
            existing_audio.spid,
            get_mp3_duration_ms(mp3_path),
        ):
            slide_file.write_bytes(
                ET.tostring(slide_root, encoding="UTF-8", xml_declaration=True)
            )

        return

    add_audio_to_slide(work_dir, slide_path, mp3_path)
