"""TTS Manager for bridging TTS providers with QML UI."""

from pathlib import Path

from PySide6.QtCore import (
    Property,
    QObject,
    QSettings,
    QStandardPaths,
    QThreadPool,
    Signal,
    Slot,
)
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
from PySide6.QtQml import QmlElement, QmlSingleton

from slide_voice_app.qml_modules.SlideVoiceApp.workers import (
    AudioGenerateWorker,
    VoiceFetchWorker,
)
from slide_voice_app.tts.google import GoogleTTSProvider
from slide_voice_app.tts.provider import ProviderInfo, TTSProvider, Voice

QML_IMPORT_NAME = "SlideVoiceApp"
QML_IMPORT_MAJOR_VERSION = 1


PROVIDER_REGISTRY: list[type[TTSProvider]] = [
    GoogleTTSProvider,
]


@QmlElement
@QmlSingleton
class TTSManager(QObject):
    """Manages TTS providers and audio playback for the QML UI."""

    # list[Voice]
    voicesReady = Signal(list)
    # path to generated file
    generationFinished = Signal(str)
    # error message
    errorOccurred = Signal(str)

    currentSlideHasAudioChanged = Signal()
    isGeneratingChanged = Signal()
    isFetchingVoicesChanged = Signal()
    isPlayingChanged = Signal()
    currentProviderChanged = Signal()

    def __init__(self, parent: QObject | None = None):
        super().__init__(parent)

        self._providers: dict[str, TTSProvider] = {}
        self._current_provider_id: str = ""
        self._provider_info: dict[str, ProviderInfo] = {}
        self._provider_classes: dict[str, type[TTSProvider]] = {}

        for provider_class in PROVIDER_REGISTRY:
            info = provider_class.get_provider_info()
            self._provider_info[info.id] = info
            self._provider_classes[info.id] = provider_class

        self._voices: list[Voice] = []
        self._is_fetching_voices: bool = False
        self._is_generating: bool = False
        self._current_slide_has_audio: bool = False

        self._thread_pool = QThreadPool.globalInstance()

        self._media_player = QMediaPlayer(self)
        self._media_player.setAudioOutput(QAudioOutput(self))
        self._media_player.playingChanged.connect(self.isPlayingChanged)
        self._media_player.errorOccurred.connect(self._on_media_error)

        self._settings = QSettings()

    def _settings_key(self, provider_id: str, setting_key: str) -> str:
        """Get the QSettings key for a provider setting."""
        return f"{provider_id}/{setting_key}"

    def _is_provider(self, provider_id: str) -> bool:
        """Check if the given provider ID is the current provider. Emit error if not."""
        exists = provider_id in self._provider_info

        if not exists:
            self.errorOccurred.emit(f"Unknown provider ID: {provider_id}")

        return exists

    def _is_created(self, provider_id: str) -> bool:
        """Check if the provider instance has been created."""
        return provider_id in self._providers

    def _get_output_file_path(self) -> str:
        """Get the path to the generated audio output file."""
        temp_location = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.TempLocation
        )
        temp_dir = Path(temp_location) / "slide-voice-app"
        temp_dir.mkdir(parents=True, exist_ok=True)
        return str(temp_dir / "slide-voice-app.mp3")

    @Property(str, constant=True)
    def outputFile(self) -> str:
        """The path to the generated audio output file."""
        return self._get_output_file_path()

    @Property(bool, notify=isGeneratingChanged)
    def isGenerating(self) -> bool:
        """Whether audio is currently being generated."""
        return self._is_generating

    @Property(bool, notify=isFetchingVoicesChanged)
    def isFetchingVoices(self) -> bool:
        """Whether voices are currently being fetched."""
        return self._is_fetching_voices

    @Property(bool, notify=isPlayingChanged)
    def isPlaying(self) -> bool:
        """Whether audio is currently playing."""
        return self._media_player.isPlaying()

    @Property(str, notify=currentProviderChanged)
    def currentProvider(self) -> str:
        """The currently selected provider ID."""
        return self._current_provider_id

    def get_currentSlideHasAudio(self) -> bool:
        """Whether the current slide has generated audio ready to save."""
        return self._current_slide_has_audio

    def set_currentSlideHasAudio(self, value: bool):
        if self._current_slide_has_audio == value:
            return

        self._current_slide_has_audio = value
        self.currentSlideHasAudioChanged.emit()

    currentSlideHasAudio = Property(
        bool,
        get_currentSlideHasAudio,
        set_currentSlideHasAudio,
        notify=currentSlideHasAudioChanged,
    )

    @Slot(result=list)
    def getAvailableProviders(self) -> list[dict[str, str]]:
        """Get list of available providers with their info.

        Returns:
            List of dicts with 'id' and 'name' keys.
        """
        return [
            {
                "id": info.id,
                "name": info.name,
            }
            for info in self._provider_info.values()
        ]

    @Slot(str, result=list)
    def getProviderSettings(self, provider_id: str) -> list[dict[str, object]]:
        """Get the settings schema for a provider.

        Args:
            provider_id: The provider ID.

        Returns:
            List of setting definitions as dicts for QML consumption.
        """
        if not self._is_provider(provider_id):
            return []

        info = self._provider_info.get(provider_id)
        assert info is not None

        settings = []

        for setting in info.settings:
            settings_key = self._settings_key(provider_id, setting.key)
            settings.append(
                {
                    "key": settings_key,
                    "label": setting.label,
                    "type": setting.setting_type.value,
                    "placeholder": setting.placeholder,
                    "value": self._settings.value(settings_key, ""),
                }
            )

        return settings

    def _get_provider_setting_values(self, provider_id: str) -> dict[str, str]:
        """Get all saved setting values for a provider.

        Args:
            provider_id: The provider ID.

        Returns:
            Dictionary mapping setting keys to their saved values.
        """
        if not self._is_provider(provider_id):
            return {}

        info = self._provider_info.get(provider_id)
        assert info is not None

        return {
            s.key: str(self._settings.value(self._settings_key(provider_id, s.key), ""))
            for s in info.settings
        }

    @Slot(str)
    def setProvider(self, provider_id: str):
        """Set the current TTS provider and configure it with saved settings.

        Args:
            provider_id: ID of the provider
        """
        if not self._is_provider(provider_id):
            return

        if not self._is_created(provider_id):
            provider_class = self._provider_classes[provider_id]
            self._providers[provider_id] = provider_class()

        provider = self._providers[provider_id]

        settings = self._get_provider_setting_values(provider_id)
        provider.configure(settings)
        self._current_provider_id = provider_id
        self.currentProviderChanged.emit()
        self.fetchVoices()

    def _on_voices_fetched(self, voices: list[Voice]):
        """Handle successful voice fetch."""
        self._voices = voices
        self._is_fetching_voices = False
        self.isFetchingVoicesChanged.emit()
        self.voicesReady.emit(
            [
                {
                    "id": v.id,
                    "name": v.name,
                    "languageCode": v.language_code,
                    "gender": v.gender,
                }
                for v in voices
            ]
        )

    def _on_voices_error(self, error_msg: str):
        """Handle voice fetch error."""
        self._is_fetching_voices = False
        self.isFetchingVoicesChanged.emit()
        self.voicesReady.emit([])
        self.errorOccurred.emit(f"Failed to fetch voices: {error_msg}")

    @Slot()
    def fetchVoices(self):
        """Fetch available voices from the current provider."""
        if not self._current_provider_id:
            self.voicesReady.emit([])
            return

        if self._is_fetching_voices:
            return

        self._is_fetching_voices = True
        self.isFetchingVoicesChanged.emit()
        provider = self._providers[self._current_provider_id]
        worker = VoiceFetchWorker(provider)
        worker.signals.finished.connect(self._on_voices_fetched)
        worker.signals.error.connect(self._on_voices_error)
        self._thread_pool.start(worker)

    @Slot(str, str)
    def generateAudio(self, text: str, voice_id: str, language_code: str):
        """Generate audio from text.

        Args:
            text: The text to convert to speech.
            voice_id: The ID of the voice to use.
            language_code: The language code for the voice.
        """
        if not self._current_provider_id:
            self.errorOccurred.emit("No provider configured")
            return

        if not text.strip():
            self.errorOccurred.emit("No text to generate audio from")
            return

        if not voice_id:
            self.errorOccurred.emit("No voice selected")
            return

        if self._is_generating:
            return

        self._is_generating = True
        self.isGeneratingChanged.emit()
        output_path = Path(self._get_output_file_path())
        provider = self._providers[self._current_provider_id]
        worker = AudioGenerateWorker(
            provider, text, voice_id, language_code, output_path
        )
        worker.signals.finished.connect(self._on_audio_generated)
        worker.signals.error.connect(self._on_audio_error)
        self._thread_pool.start(worker)

    def _on_audio_generated(self, file_path: str):
        """Handle successful audio generation."""
        self._is_generating = False
        self.isGeneratingChanged.emit()
        self.currentSlideHasAudio = True
        self.generationFinished.emit(file_path)

    def _on_audio_error(self, error_msg: str):
        """Handle audio generation error."""
        self._is_generating = False
        self.isGeneratingChanged.emit()
        self.errorOccurred.emit(f"Failed to generate audio: {error_msg}")

    @Slot(str)
    def playAudio(self, file_path: str):
        """Play an audio file.

        Args:
            file_path: Path to the audio file to play.
        """
        if not file_path:
            self.errorOccurred.emit("No audio file to play")
            return

        # Clear buffer
        self._media_player.setSource("")
        self._media_player.setSource(file_path)
        self._media_player.play()

    @Slot()
    def stopAudio(self):
        """Stop audio playback."""
        self._media_player.stop()

    @Slot(str, str, str)
    def generateAndPlay(self, text: str, voice_id: str, language_code: str):
        """Generate audio and play it when ready.

        Args:
            text: The text to convert to speech.
            voice_id: The ID of the voice to use.
            language_code: The language code for the voice.
        """

        def play_on_finish(path: str):
            self.generationFinished.disconnect(play_on_finish)
            self.playAudio(path)

        self.generationFinished.connect(play_on_finish)
        self.generateAudio(text, voice_id, language_code)

    def _on_media_error(self, error: QMediaPlayer.Error, error_string: str):
        """Handle media player errors."""
        self.errorOccurred.emit(f"Playback error: {error_string}")
