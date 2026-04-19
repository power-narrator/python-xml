"""Google Cloud Text-to-Speech provider implementation."""

from pathlib import Path

from google.cloud import texttospeech

from power_narrator.ui.tts.provider import (
    ProviderInfo,
    SettingDefinition,
    SettingType,
    TTSProvider,
    Voice,
)
from power_narrator.ui.tts.ssml import SSMLProcessor


class GoogleTTSProvider(TTSProvider):
    """Google Cloud Text-to-Speech provider.

    Uses the google-cloud-texttospeech library to generate audio.
    Authentication can be done via:
    - API key (simple string key from Google Cloud Console)
    - Default application credentials (GOOGLE_APPLICATION_CREDENTIALS env var)
    """

    def __init__(self):
        """Initialize the Google TTS provider."""
        self._api_key: str | None = None
        self._client: texttospeech.TextToSpeechClient | None = None

    @classmethod
    def get_provider_info(cls) -> ProviderInfo:
        """Return provider info including settings schema.

        This is the single source of truth for Google Cloud TTS settings.
        """
        return ProviderInfo(
            id="google_cloud",
            name="Google Cloud",
            settings=[
                SettingDefinition(
                    "api_key",
                    "API Key",
                    SettingType.PASSWORD,
                    placeholder="Enter your Google Cloud API key...",
                ),
            ],
        )

    def configure(self, settings: dict[str, str]):
        """Configure the provider with settings.

        Args:
            settings: Dictionary with optional 'api_key' key.
        """
        api_key = settings.get("api_key", "")

        if api_key:
            self._api_key = api_key
        else:
            self._api_key = None

        self._client = None

    def _get_client(self) -> texttospeech.TextToSpeechClient:
        """Get or create the TTS client.

        Returns:
            The TextToSpeechClient instance.

        Raises:
            google.auth.exceptions.MutualTLSChannelError: If mutual TLS transport creation failed for any reason.
        """
        if self._client is None:
            if self._api_key:
                self._client = texttospeech.TextToSpeechClient(
                    client_options={"api_key": self._api_key}
                )
            else:
                # Use default application credentials
                self._client = texttospeech.TextToSpeechClient()

        return self._client

    def list_voices(self) -> list[Voice]:
        """List available voices from Google Cloud TTS.

        Returns:
            A list of Voice objects.

        Raises:
            Exception: If the API call fails.
        """
        voices: list[Voice] = []
        gender_map = {
            texttospeech.SsmlVoiceGender.MALE: "Male",
            texttospeech.SsmlVoiceGender.FEMALE: "Female",
            texttospeech.SsmlVoiceGender.NEUTRAL: "Neutral",
        }

        for voice in self._get_client().list_voices().voices:
            language_code = voice.language_codes[0] if voice.language_codes else ""
            gender = gender_map.get(voice.ssml_gender, "Unknown")
            voices.append(
                Voice(
                    voice.name,
                    voice.name,
                    language_code,
                    gender,
                )
            )

        return voices

    def generate_audio(
        self, text: str, voice_id: str, language_code: str, output_path: Path
    ) -> Path:
        """Generate audio from text using Google Cloud TTS.

        Args:
            text: The text to convert to speech.
            voice_id: The voice name/ID to use (e.g., "en-US-Wavenet-A").
            output_path: The path where the MP3 file should be saved.

        Returns:
            The path to the generated audio file.

        Raises:
            Exception: If the API call fails.
        """
        ssml_text = SSMLProcessor.to_ssml(text)
        synthesis_input = texttospeech.SynthesisInput(ssml=ssml_text)
        voice_params = texttospeech.VoiceSelectionParams(
            language_code=language_code,
            name=voice_id,
        )
        audio_config = texttospeech.AudioConfig(
            audio_encoding=texttospeech.AudioEncoding.MP3
        )
        response = self._get_client().synthesize_speech(
            input=synthesis_input,
            voice=voice_params,
            audio_config=audio_config,
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with open(output_path, "wb") as out:
            out.write(response.audio_content)

        return output_path
