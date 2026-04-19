import sys
import platform

from PySide6.QtCore import QCoreApplication, QSettings
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine

import power_narrator.ui.qml_modules.PowerNarrator
import power_narrator.ui.rc_resources  # noqa: F401


def main():
    QCoreApplication.setOrganizationName("power-narrator")
    QCoreApplication.setApplicationName("power-narrator")
    QCoreApplication.addLibraryPath("/usr/lib/qt6/plugins")
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    app = QGuiApplication(sys.argv)
    engine = QQmlApplicationEngine()

    if platform.system() == "Linux":
        QCoreApplication.addLibraryPath("/usr/lib/qt6/plugins")
        engine.addImportPath("/usr/lib/qt6/qml")

    engine.load(":/qt/qml/Main.qml")

    if not engine.rootObjects():
        sys.exit(-1)

    exit_code = app.exec()
    del engine
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
