"""Qt list models exposed to QML."""

from enum import IntEnum
from typing import Any

from PySide6.QtCore import (
    QAbstractListModel,
    QByteArray,
    QModelIndex,
    QPersistentModelIndex,
    Qt,
    Slot,
)
from PySide6.QtQml import QmlElement

from slide_voice_app.audio_identity import EMBEDDED_AUDIO_BASENAME
from slide_voice_app.tts.provider import ProviderInfo, Voice
from slide_voice_pptx.pptx_file import PptxFile

QML_IMPORT_NAME = "SlideVoiceApp"
QML_IMPORT_MAJOR_VERSION = 1


class ProviderRole(IntEnum):
    Id = Qt.ItemDataRole.UserRole + 1
    Name = Qt.ItemDataRole.UserRole + 2


@QmlElement
class ProvidersModel(QAbstractListModel):
    """Read-only model of available TTS providers."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._providers: list[ProviderInfo] = []
        self.Role = ProviderRole

    def rowCount(
        self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()
    ) -> int:
        if parent.isValid():
            return 0

        return len(self._providers)

    def data(
        self,
        index: QModelIndex | QPersistentModelIndex,
        role: int = Qt.ItemDataRole.DisplayRole,
    ) -> Any:
        if not index.isValid():
            return None

        provider = self._providers[index.row()]

        if role == self.Role.Id:
            return provider.id

        if role in {self.Role.Name, Qt.ItemDataRole.DisplayRole}:
            return provider.name

        return None

    def roleNames(self) -> dict[int, QByteArray]:
        return {
            self.Role.Id: QByteArray(b"id"),
            self.Role.Name: QByteArray(b"name"),
        }

    def setProviders(self, providers: list[ProviderInfo]) -> None:
        self.beginResetModel()
        self._providers = providers
        self.endResetModel()

    @Slot(int, result=str)
    def providerIdAt(self, row: int) -> str:
        if row < 0 or row >= len(self._providers):
            return ""

        return self._providers[row].id


class VoicesRole(IntEnum):
    Id = Qt.ItemDataRole.UserRole + 1
    Name = Qt.ItemDataRole.UserRole + 2
    LanguageCode = Qt.ItemDataRole.UserRole + 3
    Gender = Qt.ItemDataRole.UserRole + 4


@QmlElement
class VoicesModel(QAbstractListModel):
    """Read-only model of voices for the current provider."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._voices: list[Voice] = []
        self.Role = VoicesRole

    def rowCount(
        self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()
    ) -> int:
        if parent.isValid():
            return 0

        return len(self._voices)

    def data(
        self,
        index: QModelIndex | QPersistentModelIndex,
        role: int = Qt.ItemDataRole.DisplayRole,
    ) -> Any:
        if not index.isValid():
            return None

        voice = self._voices[index.row()]

        if role == self.Role.Id:
            return voice.id

        if role in {self.Role.Name, Qt.ItemDataRole.DisplayRole}:
            return voice.name

        if role == self.Role.LanguageCode:
            return voice.language_code

        if role == self.Role.Gender:
            return voice.gender

        return None

    def roleNames(self) -> dict[int, QByteArray]:
        return {
            self.Role.Id: QByteArray(b"id"),
            self.Role.Name: QByteArray(b"name"),
            self.Role.LanguageCode: QByteArray(b"languageCode"),
            self.Role.Gender: QByteArray(b"gender"),
        }

    def setVoices(self, voices: list[Voice]) -> None:
        self.beginResetModel()
        self._voices = voices
        self.endResetModel()

    def clear(self) -> None:
        self.setVoices([])

    @Slot(int, result=str)
    def languageCodeAt(self, row: int) -> str:
        if row < 0 or row >= len(self._voices):
            return ""

        return self._voices[row].language_code


class SlidesModel(QAbstractListModel):
    """Editable model of slides in the open PPTX workspace."""

    class Role(IntEnum):
        Index = Qt.ItemDataRole.UserRole + 1
        Notes = Qt.ItemDataRole.UserRole + 2
        HasEmbeddedAudio = Qt.ItemDataRole.UserRole + 3

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pptx_file: PptxFile | None = None

    def rowCount(
        self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()
    ) -> int:
        if parent.isValid() or self._pptx_file is None:
            return 0

        return len(self._pptx_file.slides)

    def data(
        self,
        index: QModelIndex | QPersistentModelIndex,
        role: int = Qt.ItemDataRole.DisplayRole,
    ) -> Any:
        slide = self.slideAt(index.row())

        if slide is None:
            return None

        if role == self.Role.Index:
            return slide.index

        if role in {
            self.Role.Notes,
            Qt.ItemDataRole.DisplayRole,
            Qt.ItemDataRole.EditRole,
        }:
            return slide.notes

        if role == self.Role.HasEmbeddedAudio:
            return any(audio.name == EMBEDDED_AUDIO_BASENAME for audio in slide.audio)

        return None

    def slideAt(self, row: int) -> Any | None:
        if self._pptx_file is None or row < 0 or row >= len(self._pptx_file.slides):
            return None

        return self._pptx_file.slides[row]

    def roleNames(self) -> dict[int, QByteArray]:
        return {
            self.Role.Index: QByteArray(b"index"),
            self.Role.Notes: QByteArray(b"notes"),
            self.Role.HasEmbeddedAudio: QByteArray(b"hasEmbeddedAudio"),
        }

    def setData(
        self,
        index: QModelIndex | QPersistentModelIndex,
        value: Any,
        role: int = Qt.ItemDataRole.EditRole,
    ) -> bool:
        slide = self.slideAt(index.row())

        if slide is None or role not in {self.Role.Notes, Qt.ItemDataRole.EditRole}:
            return False

        notes = str(value)

        slide.notes = notes
        self.dataChanged.emit(index, index)
        return True

    def flags(self, index: QModelIndex | QPersistentModelIndex) -> Qt.ItemFlag:
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags

        return (
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEditable
        )

    def setPptxFile(self, pptx_file: PptxFile | None) -> None:
        self.beginResetModel()
        self._pptx_file = pptx_file
        self.endResetModel()
