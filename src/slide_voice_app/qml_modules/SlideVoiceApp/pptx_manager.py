"""PPTX Manager for bridging PPTX file operations with QML UI."""

from pathlib import Path
from urllib.parse import urlparse
from urllib.request import url2pathname

from PySide6.QtCore import Property, QObject, Signal, Slot
from PySide6.QtQml import QmlElement, QmlSingleton

from slide_voice_app.pptx import PptxFile
from slide_voice_app.pptx.exceptions import (
    InvalidPptxError,
    RelsNotFoundError,
    SlideNotFoundError,
    SlideXmlNotFoundError,
)

QML_IMPORT_NAME = "SlideVoiceApp"
QML_IMPORT_MAJOR_VERSION = 1


@QmlElement
@QmlSingleton
class PPTXManager(QObject):
    """Manages PPTX file operations for the QML UI."""

    slidesLoaded = Signal(list)
    errorOccurred = Signal(str)
    fileLoadedChanged = Signal()

    def __init__(self, parent: QObject | None = None):
        super().__init__(parent)
        self._pptx_file: PptxFile | None = None

    @Property(bool, notify=fileLoadedChanged)
    def fileLoaded(self) -> bool:
        """Whether a PPTX file is currently loaded."""
        return self._pptx_file is not None

    def _unload_file(self):
        """Indicate file is not loaded."""
        if self._pptx_file is not None:
            self._pptx_file.close()

        self._pptx_file = None
        self.fileLoadedChanged.emit()

    @Slot(str)
    def openFile(self, file_url: str):
        """Open a PPTX file and load its slide notes.

        Args:
            file_url: string to the .pptx file.
        """
        path = Path(url2pathname(urlparse(file_url).path))

        self._unload_file()

        try:
            self._pptx_file = PptxFile.open(path)
            pptx_file = self._pptx_file
            self.fileLoadedChanged.emit()
            notes = pptx_file.get_all_slide_notes()
            slides_data = [{"notes": note} for note in notes]
            self.slidesLoaded.emit(slides_data)

        except FileNotFoundError:
            self._unload_file()
            self.errorOccurred.emit(f"File not found: {path}")
        except InvalidPptxError as e:
            self._unload_file()
            self.errorOccurred.emit(str(e))
        except RelsNotFoundError as e:
            self._unload_file()
            self.errorOccurred.emit(str(e))
        except SlideXmlNotFoundError as e:
            self._unload_file()
            self.errorOccurred.emit(str(e))
        except Exception as e:
            self._unload_file()
            self.errorOccurred.emit(f"Failed to open file: {e}")

    @Slot(int, str)
    def saveAudioForSlide(self, slide_index: int, mp3_file_url: str):
        """Insert audio into the selected slide in the loaded PPTX workspace."""
        if not mp3_file_url:
            self.errorOccurred.emit("No audio file provided")
            return

        if self._pptx_file is None:
            self.errorOccurred.emit("No PPTX file loaded")
            return

        try:
            mp3_path = Path(url2pathname(urlparse(mp3_file_url).path))
            self._pptx_file.save_audio_for_slide(slide_index, mp3_path)
        except FileNotFoundError as e:
            self.errorOccurred.emit(f"File not found: {e}")
        except SlideNotFoundError as e:
            self.errorOccurred.emit(str(e))
        except SlideXmlNotFoundError as e:
            self.errorOccurred.emit(str(e))
        except Exception as e:
            self.errorOccurred.emit(f"Failed to save audio: {e}")

    @Slot(int, str)
    def setSlideNotes(self, slide_index: int, notes: str):
        """Update in-memory notes for a slide in the loaded PPTX workspace."""
        if self._pptx_file is None:
            self.errorOccurred.emit("No PPTX file loaded")
            return

        try:
            self._pptx_file.set_slide_notes(slide_index, notes)
        except SlideNotFoundError as e:
            self.errorOccurred.emit(str(e))
        except Exception as e:
            self.errorOccurred.emit(f"Failed to update notes: {e}")

    @Slot(str)
    def exportTo(self, output_file_url: str):
        """Export the current PPTX workspace to the selected output path."""
        if not output_file_url:
            self.errorOccurred.emit("No output file provided")
            return

        if self._pptx_file is None:
            self.errorOccurred.emit("No PPTX file loaded to export")
            return
        try:
            output_path = Path(url2pathname(urlparse(output_file_url).path))

            if self._pptx_file is None:
                self.errorOccurred.emit("No PPTX file loaded to export")
                return

            self._pptx_file.export_to(output_path)
        except Exception as e:
            self.errorOccurred.emit(f"Failed to export file: {e}")
