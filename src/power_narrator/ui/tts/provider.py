"""Abstract base class for TTS providers."""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path


class SettingType(StrEnum):
    """Types of settings that providers can define."""

    STRING = "string"
    PASSWORD = "password"


@dataclass
class SettingDefinition:
    """Defines a single setting that a provider requires.

    The UI use this schema to render forms.
    """

    key: str
    label: str
    setting_type: SettingType = SettingType.STRING
    placeholder: str = ""


@dataclass
class ProviderInfo:
    """Static information about a TTS provider."""

    id: str
    name: str
    settings: list[SettingDefinition]


@dataclass
class Voice:
    """Represents a voice available from a TTS provider."""

    id: str
    name: str
    language_code: str
    gender: str


class TTSProvider(ABC):
    """Abstract base class for TTS providers."""

    @classmethod
    @abstractmethod
    def get_provider_info(cls) -> ProviderInfo:
        """Return static information about this provider including settings schema.

        Called without instantiating the provider.

        Returns:
            ProviderInfo containing the provider's metadata and settings schema.
        """
        pass

    @abstractmethod
    def configure(self, settings: dict[str, str]):
        """Configure the provider with the given settings.

        Args:
            settings: Dictionary mapping setting keys to their values.
                Keys should match those defined in get_provider_info().settings.

        Raises:
            ValueError: If required settings are missing or invalid.
        """
        pass

    @abstractmethod
    def list_voices(self) -> list[Voice]:
        """List available voices from this provider.

        This method may make a network request hence blocks until complete.
        Call from a background thread.

        Returns:
            A list of Voice objects available from this provider.

        Raises:
            Exception: If the API call fails or credentials are invalid.
        """
        pass

    @abstractmethod
    def generate_audio(
        self, text: str, voice_id: str, language_code: str, output_path: Path
    ) -> Path:
        """Generate audio from text and save to a file.

        This method may make a network request hence blocks until complete.
        Call from a background thread.

        Args:
            text: The text to convert to speech.
            voice_id: The ID of the voice to use.
            language_code: The language code for the voice.
            output_path: The path where the audio file should be saved.

        Returns:
            The path to the generated audio file.

        Raises:
            Exception: If the API call fails or credentials are invalid.
        """
        pass
