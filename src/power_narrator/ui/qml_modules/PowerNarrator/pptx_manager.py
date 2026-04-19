"""PPTX Manager for bridging PPTX file operations with QML UI."""

from pathlib import Path
from urllib.parse import urlparse
from urllib.request import url2pathname

from PySide6.QtCore import Property, QObject, Signal, Slot
from PySide6.QtQml import QmlElement, QmlSingleton

from power_narrator.pptx import PptxFile
from power_narrator.pptx.exceptions import (
    AudioNotFoundError,
    InvalidPptxError,
    RelsNotFoundError,
    SlideNotFoundError,
    SlideXmlNotFoundError,
)
from power_narrator.ui.audio_identity import EMBEDDED_AUDIO_BASENAME
from power_narrator.ui.qml_modules.PowerNarrator.models import SlidesModel

QML_IMPORT_NAME = "PowerNarrator"
QML_IMPORT_MAJOR_VERSION = 1


@QmlElement
@QmlSingleton
class PPTXManager(QObject):
    """Manages PPTX file operations for the QML UI."""

    errorOccurred = Signal(str)
    fileLoadedChanged = Signal()
    currentSlideIndexChanged = Signal()
    currentSlideNotesChanged = Signal()
    currentSlideHasEmbeddedAudioChanged = Signal()

    def __init__(self, parent: QObject | None = None):
        super().__init__(parent)
        self._pptx_file: PptxFile | None = None
        self._current_slide_index = -1
        self._slides_model = SlidesModel(self)
        self._slides_model.modelReset.connect(self._emit_current_slide_state)
        self._slides_model.dataChanged.connect(self._on_slides_data_changed)

    @Property(bool, notify=fileLoadedChanged)
    def fileLoaded(self) -> bool:
        """Whether a PPTX file is currently loaded."""
        return self._pptx_file is not None

    @Property(QObject, constant=True)
    def slidesModel(self) -> QObject:
        """Model of slides in the currently loaded PPTX file."""
        return self._slides_model

    def _normalize_slide_index(self, value: int) -> int:
        if self._pptx_file is None:
            return -1

        slide_count = len(self._pptx_file.slides)

        if slide_count == 0:
            return -1

        if value < 0 or value >= slide_count:
            return -1

        return value

    def _emit_current_slide_state(self) -> None:
        self.currentSlideNotesChanged.emit()
        self.currentSlideHasEmbeddedAudioChanged.emit()

    def _on_slides_data_changed(self, top_left, bottom_right) -> None:
        if top_left.row() == self._current_slide_index:
            self._emit_current_slide_state()

    def _unload_file(self) -> None:
        """Indicate file is not loaded."""
        if self._pptx_file is not None:
            self._pptx_file.close()

        self._pptx_file = None
        self._slides_model.setPptxFile(None)
        self.fileLoadedChanged.emit()
        self.setCurrentSlideIndex(-1)

    def getCurrentSlideIndex(self) -> int:
        """Index of the currently selected slide."""
        return self._current_slide_index

    def setCurrentSlideIndex(self, value: int) -> None:
        self._current_slide_index = self._normalize_slide_index(value)
        self.currentSlideIndexChanged.emit()
        self._emit_current_slide_state()

    currentSlideIndex = Property(
        int,
        getCurrentSlideIndex,
        setCurrentSlideIndex,
        notify=currentSlideIndexChanged,
    )

    def getCurrentSlideNotes(self) -> str:
        """Notes for the currently selected slide."""
        slide = self._slides_model.slideAt(self._current_slide_index)
        return "" if slide is None else slide.notes

    def setCurrentSlideNotes(self, notes: str) -> None:
        index = self._slides_model.index(self._current_slide_index, 0)
        self._slides_model.setData(index, notes, SlidesModel.Role.Notes)

    currentSlideNotes = Property(
        str,
        getCurrentSlideNotes,
        setCurrentSlideNotes,
        notify=currentSlideNotesChanged,
    )

    @Property(bool, notify=currentSlideHasEmbeddedAudioChanged)
    def currentSlideHasEmbeddedAudio(self) -> bool:
        index = self._slides_model.index(self._current_slide_index, 0)
        value = self._slides_model.data(index, SlidesModel.Role.HasEmbeddedAudio)
        return bool(value)

    @Slot(str)
    def openFile(self, file_url: str):
        """Open a PPTX file and load its slide notes."""
        path = Path(url2pathname(urlparse(file_url).path))

        self._unload_file()

        try:
            self._pptx_file = PptxFile.open(path)
            self._slides_model.setPptxFile(self._pptx_file)
            self.fileLoadedChanged.emit()
            self.setCurrentSlideIndex(0)
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

    @Slot(str)
    def saveAudioForCurrentSlide(self, mp3_file_url: str):
        """Insert audio into the selected slide in the loaded PPTX workspace."""
        if not mp3_file_url:
            self.errorOccurred.emit("No audio file provided")
            return

        if self._pptx_file is None:
            self.errorOccurred.emit("No PPTX file loaded")
            return

        if self._current_slide_index < 0:
            self.errorOccurred.emit("No slide selected")
            return

        try:
            mp3_path = Path(url2pathname(urlparse(mp3_file_url).path))
            self._pptx_file.save_audio_for_slide(self._current_slide_index, mp3_path)
            self._emit_current_slide_state()
        except FileNotFoundError as e:
            self.errorOccurred.emit(f"File not found: {e}")
        except SlideNotFoundError as e:
            self.errorOccurred.emit(str(e))
        except SlideXmlNotFoundError as e:
            self.errorOccurred.emit(str(e))
        except Exception as e:
            self.errorOccurred.emit(f"Failed to save audio: {e}")

    @Slot()
    def deleteAudioForCurrentSlide(self):
        """Delete app-managed audio from the selected slide."""
        if self._pptx_file is None:
            self.errorOccurred.emit("No PPTX file loaded")
            return

        if self._current_slide_index < 0:
            self.errorOccurred.emit("No slide selected")
            return

        try:
            self._pptx_file.delete_audio_for_slide(
                self._current_slide_index, EMBEDDED_AUDIO_BASENAME
            )
            self._emit_current_slide_state()
        except AudioNotFoundError as e:
            self.errorOccurred.emit(str(e))
        except SlideNotFoundError as e:
            self.errorOccurred.emit(str(e))
        except SlideXmlNotFoundError as e:
            self.errorOccurred.emit(str(e))
        except Exception as e:
            self.errorOccurred.emit(f"Failed to delete audio: {e}")

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
            self._pptx_file.export_to(output_path)
        except Exception as e:
            self.errorOccurred.emit(f"Failed to export file: {e}")
