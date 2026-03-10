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

from slide_voice_app.audio_identity import EMBEDDED_AUDIO_FILENAME
from slide_voice_app.qml_modules.SlideVoiceApp.models import ProvidersModel, VoicesModel
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

    errorOccurred = Signal(str)
    hasGeneratedAudioChanged = Signal()
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

        self._is_fetching_voices = False
        self._is_generating = False
        self._has_generated_audio = False

        self._providers_model = ProvidersModel(self)
        self._providers_model.setProviders(list(self._provider_info.values()))
        self._voices_model = VoicesModel(self)

        self._thread_pool = QThreadPool.globalInstance()

        self._media_player = QMediaPlayer(self)
        self._media_player.setAudioOutput(QAudioOutput(self))
        self._media_player.playingChanged.connect(self.isPlayingChanged)
        self._media_player.errorOccurred.connect(self._on_media_error)

        self._settings = QSettings()

        if self._provider_info:
            first_provider_id = next(iter(self._provider_info))
            self.setCurrentProvider(first_provider_id)

    def _settings_key(self, provider_id: str, setting_key: str) -> str:
        """Get the QSettings key for a provider setting."""
        return f"{provider_id}/{setting_key}"

    def _is_provider(self, provider_id: str) -> bool:
        """Check if the given provider ID exists. Emit error if not."""
        exists = provider_id in self._provider_info

        if not exists:
            self.errorOccurred.emit(f"Unknown provider ID: {provider_id}")

        return exists

    def _provider_for(self, provider_id: str) -> TTSProvider:
        if provider_id not in self._providers:
            provider_class = self._provider_classes[provider_id]
            self._providers[provider_id] = provider_class()

        return self._providers[provider_id]

    def _get_output_file_path(self) -> str:
        """Get the path to the generated audio output file."""
        temp_location = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.TempLocation
        )
        temp_dir = Path(temp_location) / "slide-voice-app"
        temp_dir.mkdir(parents=True, exist_ok=True)
        return str(temp_dir / EMBEDDED_AUDIO_FILENAME)

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

    def getHasGeneratedAudio(self) -> bool:
        """Whether generated audio is ready to insert into a slide."""
        return self._has_generated_audio

    def setHasGeneratedAudio(self, value: bool) -> None:
        if self._has_generated_audio == value:
            return

        self._has_generated_audio = value
        self.hasGeneratedAudioChanged.emit()

    hasGeneratedAudio = Property(
        bool,
        getHasGeneratedAudio,
        setHasGeneratedAudio,
        notify=hasGeneratedAudioChanged,
    )

    @Property(ProvidersModel, constant=True)
    def providersModel(self) -> QObject:
        """Model of available providers."""
        return self._providers_model

    @Property(VoicesModel, constant=True)
    def voicesModel(self) -> QObject:
        """Model of voices for the current provider."""
        return self._voices_model

    def getCurrentProvider(self) -> str:
        """The currently selected provider ID."""
        return self._current_provider_id

    def _get_provider_setting_values(self, provider_id: str) -> dict[str, str]:
        """Get all saved setting values for a provider."""
        if not self._is_provider(provider_id):
            return {}

        info = self._provider_info.get(provider_id)
        assert info is not None

        return {
            s.key: str(self._settings.value(self._settings_key(provider_id, s.key), ""))
            for s in info.settings
        }

    def setCurrentProvider(self, provider_id: str):
        """Set or reconfigure the current TTS provider."""
        if not self._is_provider(provider_id):
            return

        provider = self._provider_for(provider_id)
        settings = self._get_provider_setting_values(provider_id)
        provider.configure(settings)

        changed = self._current_provider_id != provider_id
        self._current_provider_id = provider_id

        if changed:
            self.currentProviderChanged.emit()

        self.fetchVoices()

    currentProvider = Property(
        str,
        getCurrentProvider,
        setCurrentProvider,
        notify=currentProviderChanged,
    )

    @Slot(str, result=list)
    def getProviderSettings(self, provider_id: str) -> list[dict[str, object]]:
        """Get the settings schema for a provider."""
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

    def _on_voices_fetched(self, voices: list[Voice]):
        """Handle successful voice fetch."""
        self._is_fetching_voices = False
        self._voices_model.setVoices(voices)
        self.isFetchingVoicesChanged.emit()

    def _on_voices_error(self, error_msg: str):
        """Handle voice fetch error."""
        self._is_fetching_voices = False
        self._voices_model.clear()
        self.isFetchingVoicesChanged.emit()
        self.errorOccurred.emit(f"Failed to fetch voices: {error_msg}")

    @Slot()
    def fetchVoices(self):
        """Fetch available voices from the current provider."""
        if not self._current_provider_id:
            self._voices_model.clear()
            return

        if self._is_fetching_voices:
            return

        self._voices_model.clear()
        self._is_fetching_voices = True
        self.isFetchingVoicesChanged.emit()
        provider = self._provider_for(self._current_provider_id)
        worker = VoiceFetchWorker(provider)
        worker.signals.finished.connect(self._on_voices_fetched)
        worker.signals.error.connect(self._on_voices_error)
        self._thread_pool.start(worker)

    def _can_generate_audio(self, text: str, voice_id: str) -> bool:
        if not self._current_provider_id:
            self.errorOccurred.emit("No provider configured")
            return False

        if not text.strip():
            self.errorOccurred.emit("No text to generate audio from")
            return False

        if not voice_id:
            self.errorOccurred.emit("No voice selected")
            return False

        if self._is_generating:
            return False

        return True

    def _start_generation(self, text: str, voice_id: str, language_code: str) -> bool:
        if not self._can_generate_audio(text, voice_id):
            return False

        self.setHasGeneratedAudio(False)
        self._is_generating = True
        self.isGeneratingChanged.emit()
        output_path = Path(self._get_output_file_path())
        provider = self._provider_for(self._current_provider_id)
        worker = AudioGenerateWorker(
            provider, text, voice_id, language_code, output_path
        )
        worker.signals.finished.connect(self._on_audio_generated)
        worker.signals.error.connect(self._on_audio_error)
        self._thread_pool.start(worker)
        return True

    @Slot(str, str, str, result=bool)
    def generateAudio(self, text: str, voice_id: str, language_code: str) -> bool:
        """Generate audio from text."""
        return self._start_generation(text, voice_id, language_code)

    def _on_audio_generated(self, file_path: str):
        """Handle successful audio generation."""
        self._is_generating = False
        self.isGeneratingChanged.emit()
        self.setHasGeneratedAudio(True)
        self.playAudio(file_path)

    def _on_audio_error(self, error_msg: str):
        """Handle audio generation error."""
        self._is_generating = False
        self.isGeneratingChanged.emit()
        self.errorOccurred.emit(f"Failed to generate audio: {error_msg}")

    @Slot(str)
    def playAudio(self, file_path: str):
        """Play an audio file."""
        if not file_path:
            self.errorOccurred.emit("No audio file to play")
            return

        self._media_player.setSource("")
        self._media_player.setSource(file_path)
        self._media_player.play()

    @Slot()
    def stopAudio(self):
        """Stop audio playback."""
        self._media_player.stop()

    def _on_media_error(self, error: QMediaPlayer.Error, error_string: str):
        """Handle media player errors."""
        self.errorOccurred.emit(f"Playback error: {error_string}")
