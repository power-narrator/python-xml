import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtCore

import SlideVoiceApp

ApplicationWindow {
    id: settingsWindow
    title: "Settings"
    width: 500
    height: 400
    modality: Qt.ApplicationModal
    flags: Qt.Dialog

    ListModel {
        id: settingsModel
    }

    Settings {
        id: appSettings
    }

    function loadProviderSettings(providerId) {
        settingsModel.clear();
        TTSManager.getProviderSettings(providerId).forEach(setting => settingsModel.append(setting));
    }

    function saveSettings() {
        if (providerComboBox.currentIndex < 0)
            return;

        // currentValue is not reliable on first load for this model,
        // so resolve the provider ID from the current index instead.
        let providerId = TTSManager.providersModel.providerIdAt(providerComboBox.currentIndex);

        for (let i = 0; i < settingsModel.count; i++) {
            let setting = settingsModel.get(i);
            appSettings.setValue(setting.key, setting.value);
        }

        TTSManager.currentProvider = providerId;
    }

    Component.onCompleted: {
        providerComboBox.currentIndex = providerComboBox.indexOfValue(TTSManager.currentProvider);

        if (providerComboBox.currentIndex >= 0) {
            loadProviderSettings(TTSManager.providersModel.providerIdAt(providerComboBox.currentIndex));
        }
    }

    Connections {
        target: TTSManager

        function onCurrentProviderChanged() {
            providerComboBox.currentIndex = providerComboBox.indexOfValue(TTSManager.currentProvider);
        }
    }

    ColumnLayout {
        anchors.fill: parent
        anchors.margins: 20
        spacing: 20

        RowLayout {
            Layout.fillWidth: true
            spacing: 10

            Label {
                text: "Provider:"
            }

            ComboBox {
                id: providerComboBox
                Layout.fillWidth: true
                model: TTSManager.providersModel
                textRole: "name"
                valueRole: "id"

                onCurrentIndexChanged: {
                    if (currentIndex >= 0) {
                        // currentValue is not reliable on first load for this model,
                        // so resolve the provider ID from the current index instead.
                        settingsWindow.loadProviderSettings(TTSManager.providersModel.providerIdAt(currentIndex));
                    }
                }
            }
        }

        ListView {
            Layout.fillWidth: true
            Layout.fillHeight: true
            spacing: 15
            model: settingsModel

            delegate: RowLayout {
                id: settingList
                required property string label
                required property string value
                required property string type
                required property string placeholder
                required property var model

                width: parent.width
                spacing: 5

                Label {
                    text: settingList.label
                }

                TextField {
                    Layout.fillWidth: true
                    text: settingList.value
                    placeholderText: settingList.placeholder
                    echoMode: settingList.type === "password" ? TextInput.Password : TextInput.Normal

                    onTextChanged: settingList.ListView.view.model.setProperty(settingList.model.index, "value", text)
                }
            }
        }

        RowLayout {
            Layout.alignment: Qt.AlignRight
            spacing: 10

            Button {
                text: "Cancel"
                onClicked: settingsWindow.close()
            }

            Button {
                text: "Save"
                highlighted: true
                onClicked: {
                    settingsWindow.saveSettings();
                    settingsWindow.close();
                }
            }
        }
    }
}
