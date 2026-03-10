import QtQuick
import QtQuick.Controls
import QtQuick.Dialogs
import QtQuick.Layouts

import SlideVoiceApp

ApplicationWindow {
    id: window
    width: 1024
    height: 600
    visible: true
    title: "Slide Voice App"

    menuBar: MenuBar {
        Menu {
            title: "File"

            Action {
                text: "Open"
                onTriggered: fileDialog.open()
            }

            Action {
                text: "Save"
                enabled: PPTXManager.fileLoaded
                onTriggered: saveDialog.open()
            }

            Action {
                text: "Settings"
                onTriggered: settingsLoader.active = true
            }
        }
    }

    FileDialog {
        id: fileDialog
        title: "Open PowerPoint File"
        nameFilters: ["PowerPoint files (*.pptx)"]
        onAccepted: PPTXManager.openFile(selectedFile)
    }

    FileDialog {
        id: saveDialog
        title: "Export PowerPoint File"
        fileMode: FileDialog.SaveFile
        nameFilters: ["PowerPoint files (*.pptx)"]
        onAccepted: PPTXManager.exportTo(selectedFile)
    }

    Loader {
        id: settingsLoader
        active: false
        source: "SettingsWindow.qml"

        onLoaded: item.visible = true
    }

    Connections {
        target: settingsLoader.item

        function onClosing() {
            settingsLoader.active = false;
        }
    }

    Connections {
        target: PPTXManager

        function onErrorOccurred(message) {
            errorDialog.text = message;
            errorDialog.open();
        }

        function onCurrentSlideIndexChanged() {
            if (slideList.currentIndex !== PPTXManager.currentSlideIndex) {
                slideList.currentIndex = PPTXManager.currentSlideIndex;
            }
        }

        function onCurrentSlideNotesChanged() {
            if (!notesEditor.activeFocus || notesEditor.text !== PPTXManager.currentSlideNotes) {
                notesEditor.text = PPTXManager.currentSlideNotes;
            }
        }
    }

    Connections {
        target: TTSManager

        function onErrorOccurred(message) {
            errorDialog.text = message;
            errorDialog.open();
        }

        function onCurrentProviderChanged() {
            providerComboBox.currentIndex = providerComboBox.indexOfValue(TTSManager.currentProvider);
        }
    }

    Connections {
        // Property constant attribute does not work, model has QObject in inheritance chain
        target: TTSManager.voicesModel // qmllint disable stale-property-read incompatible-type

        function onModelReset() {
            voiceComboBox.currentIndex = voiceComboBox.count > 0 ? 0 : -1;
        }
    }

    MessageDialog {
        id: errorDialog
        title: "Error"
        buttons: MessageDialog.Ok
    }

    Component.onCompleted: {
        providerComboBox.currentIndex = providerComboBox.indexOfValue(TTSManager.currentProvider);
        notesEditor.text = PPTXManager.currentSlideNotes;
    }

    RowLayout {
        anchors.fill: parent
        anchors.margins: 10
        spacing: 10

        ListView {
            id: slideList
            Layout.preferredWidth: 140
            Layout.fillHeight: true
            model: PPTXManager.slidesModel
            spacing: 10

            delegate: Rectangle {
                id: slideItem
                required property var model

                width: ListView.view.width
                height: 40
                radius: 4
                border.width: ListView.isCurrentItem ? 3 : 0
                border.color: palette.highlight

                Text {
                    anchors.centerIn: parent
                    text: `Slide ${slideItem.model.index + 1}`
                }

                MouseArea {
                    anchors.fill: parent
                    onClicked: slideItem.ListView.view.currentIndex = slideItem.model.index
                }
            }

            onCurrentIndexChanged: {
                PPTXManager.currentSlideIndex = currentIndex;
                TTSManager.hasGeneratedAudio = false;
            }
        }

        ColumnLayout {
            spacing: 10

            ScrollView {
                Layout.fillWidth: true
                Layout.fillHeight: true

                TextArea {
                    id: notesEditor
                    placeholderText: "Slide notes..."
                    wrapMode: TextArea.Wrap
                    persistentSelection: true
                    readonly property string textToGenerate: selectedText.length > 0 ? selectedText : text

                    onTextChanged: {
                        PPTXManager.currentSlideNotes = text;
                    }
                }
            }

            RowLayout {
                Layout.fillWidth: true
                spacing: 10

                ComboBox {
                    id: providerComboBox
                    // Property constant attribute does not work
                    model: TTSManager.providersModel // qmllint disable stale-property-read
                    textRole: "name"
                    valueRole: "id"
                    displayText: currentIndex >= 0 ? currentText : "Provider"
                    enabled: !TTSManager.isGenerating && !TTSManager.isFetchingVoices

                    // currentValue is not reliable on first load for this model,
                    // so resolve the provider ID from the current index instead.
                    onActivated: TTSManager.currentProvider = TTSManager.providersModel.providerIdAt(currentIndex)
                }

                ComboBox {
                    id: voiceComboBox
                    Layout.preferredWidth: 300
                    model: TTSManager.voicesModel
                    valueRole: "id"
                    textRole: "name"
                    enabled: !TTSManager.isGenerating && !TTSManager.isFetchingVoices

                    displayText: {
                        if (TTSManager.isFetchingVoices) {
                            return "Loading...";
                        }

                        if (currentIndex >= 0) {
                            return currentText;
                        }

                        return "Voice";
                    }

                    delegate: ItemDelegate {
                        required property string name
                        required property string languageCode
                        required property string gender

                        width: ListView.view.width
                        text: `${name} (${languageCode}, ${gender})`
                    }
                }

                Item {
                    Layout.fillWidth: true
                }

                Button {
                    text: TTSManager.isPlaying ? "Stop" : "Preview"
                    enabled: !TTSManager.isGenerating && !TTSManager.isFetchingVoices && voiceComboBox.currentIndex >= 0 && notesEditor.textToGenerate.trim().length > 0

                    onClicked: {
                        if (TTSManager.isPlaying) {
                            TTSManager.stopAudio();
                        } else {
                            TTSManager.generateAudio(notesEditor.textToGenerate, voiceComboBox.currentValue, TTSManager.voicesModel.languageCodeAt(voiceComboBox.currentIndex));
                        }
                    }
                }

                Button {
                    text: "Insert Audio"
                    enabled: !TTSManager.isGenerating && PPTXManager.fileLoaded && TTSManager.hasGeneratedAudio
                    onClicked: PPTXManager.saveAudioForCurrentSlide(TTSManager.outputFile)
                }

                Button {
                    text: "Delete Audio"
                    enabled: !TTSManager.isGenerating && PPTXManager.fileLoaded && PPTXManager.currentSlideHasEmbeddedAudio
                    onClicked: PPTXManager.deleteAudioForCurrentSlide()
                }
            }
        }
    }

    footer: RowLayout {
        BusyIndicator {
            running: TTSManager.isGenerating || TTSManager.isFetchingVoices
            Layout.preferredHeight: 20
            Layout.preferredWidth: 20
        }

        Label {
            Layout.fillWidth: true
            text: {
                if (TTSManager.isGenerating) {
                    return "Generating audio...";
                }

                if (TTSManager.isFetchingVoices) {
                    return "Fetching voices...";
                }

                if (TTSManager.isPlaying) {
                    return "Playing audio...";
                }

                return "";
            }
        }
    }
}
