"""TTS Workers for background tasks."""

from pathlib import Path

from PySide6.QtCore import QObject, QRunnable, Signal

from power_narrator.ui.tts.provider import TTSProvider


class BaseWorkerSignals(QObject):
    """Base signals for workers."""

    error = Signal(str)


class BaseWorker(QRunnable):
    """Base worker handling error capture and signal emission."""

    def __init__(self, signals: BaseWorkerSignals):
        super().__init__()
        self.signals = signals

    def run(self):
        """Run the worker logic with error handling."""
        try:
            self.work()
        except Exception as e:
            self.signals.error.emit(str(e))

    def work(self):
        """Actual worker logic."""
        raise NotImplementedError


class VoiceFetchWorkerSignals(BaseWorkerSignals):
    """Signals for VoiceFetchWorker."""

    # list[Voice]
    finished = Signal(list)


class VoiceFetchWorker(BaseWorker):
    """Worker to fetch voices in a background thread."""

    def __init__(self, provider: TTSProvider):
        signals = VoiceFetchWorkerSignals()
        super().__init__(signals)
        self.provider = provider
        self.signals: VoiceFetchWorkerSignals = signals

    def work(self):
        """Fetch voices from the provider."""
        voices = self.provider.list_voices()
        self.signals.finished.emit(voices)


class AudioGenerateWorkerSignals(BaseWorkerSignals):
    """Signals for AudioGenerateWorker."""

    # path to generated file
    finished = Signal(str)


class AudioGenerateWorker(BaseWorker):
    """Worker to generate audio in a background thread."""

    def __init__(
        self,
        provider: TTSProvider,
        text: str,
        voice_id: str,
        language_code: str,
        output_path: Path,
    ):
        signals = AudioGenerateWorkerSignals()
        super().__init__(signals)
        self.provider = provider
        self.text = text
        self.voice_id = voice_id
        self.language_code = language_code
        self.output_path = output_path
        self.signals: AudioGenerateWorkerSignals = signals

    def work(self):
        """Generate audio from the provider."""
        result_path = self.provider.generate_audio(
            self.text, self.voice_id, self.language_code, self.output_path
        )
        self.signals.finished.emit(str(result_path))
