/****************************************************************************
** Generated QML type registration code
**
** WARNING! All changes made in this file will be lost!
*****************************************************************************/

#include <QtQml/qqml.h>
#include <QtQml/qqmlmoduleregistration.h>

#if __has_include(</Users/proglab/Documents/python-xml/src/power_narrator/ui/qml_modules/PowerNarrator/models.py>)
#  include </Users/proglab/Documents/python-xml/src/power_narrator/ui/qml_modules/PowerNarrator/models.py>
#endif
#if __has_include(</Users/proglab/Documents/python-xml/src/power_narrator/ui/qml_modules/PowerNarrator/pptx_manager.py>)
#  include </Users/proglab/Documents/python-xml/src/power_narrator/ui/qml_modules/PowerNarrator/pptx_manager.py>
#endif
#if __has_include(</Users/proglab/Documents/python-xml/src/power_narrator/ui/qml_modules/PowerNarrator/tts_manager.py>)
#  include </Users/proglab/Documents/python-xml/src/power_narrator/ui/qml_modules/PowerNarrator/tts_manager.py>
#endif


#if !defined(QT_STATIC)
#define Q_QMLTYPE_EXPORT Q_DECL_EXPORT
#else
#define Q_QMLTYPE_EXPORT
#endif
Q_QMLTYPE_EXPORT void qml_register_types_PowerNarrator()
{
    QT_WARNING_PUSH QT_WARNING_DISABLE_DEPRECATED
    qmlRegisterTypesAndRevisions<PPTXManager>("PowerNarrator", 1);
    qmlRegisterTypesAndRevisions<ProvidersModel>("PowerNarrator", 1);
    qmlRegisterTypesAndRevisions<TTSManager>("PowerNarrator", 1);
    qmlRegisterTypesAndRevisions<VoicesModel>("PowerNarrator", 1);
    QT_WARNING_POP
    qmlRegisterModule("PowerNarrator", 1, 0);
}

static const QQmlModuleRegistration powerNarratorRegistration("PowerNarrator", qml_register_types_PowerNarrator);
