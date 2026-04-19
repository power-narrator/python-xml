"""TTS provider module for Power Narrator."""

from power_narrator.ui.tts.google import GoogleTTSProvider
from power_narrator.ui.tts.provider import TTSProvider, Voice

__all__ = ["GoogleTTSProvider", "TTSProvider", "Voice"]
