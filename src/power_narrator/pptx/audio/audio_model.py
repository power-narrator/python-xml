"""Slide audio model types."""

from dataclasses import dataclass


@dataclass(slots=True)
class Audio:
    """Represents one audio entry attached to a slide."""

    name: str
    audio_rid: str
    media_rid: str
    image_rid: str
    target: str
    spid: int
