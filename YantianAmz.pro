QT       += core gui axcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++11

# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
    main.cpp \
    treadexcelthread.cpp \
    twriteexcelthread.cpp \
    yantianamz.cpp

HEADERS += \
    treadexcelthread.h \
    twriteexcelthread.h \
    yantianamz.h

FORMS += \
    yantianamz.ui

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

TRANSLATIONS += lang_Chinese.ts \
                 lang_English.ts

RESOURCES += \
    i18n_zh.qrc
