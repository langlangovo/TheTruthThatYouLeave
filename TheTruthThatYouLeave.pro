#-------------------------------------------------
#
# Project created by QtCreator 2017-04-21T07:15:29
#
#-------------------------------------------------

QT       += core gui multimedia
QT       += axcontainer
QT       += core

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = TheTruthThatYouLeave
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp \
    excel.cpp \
    excelthread.cpp \
    pieceworkdata.cpp

HEADERS  += mainwindow.h \
    excel.h \
    excelthread.h \
    pieceworkdata.h

FORMS    += mainwindow.ui

RESOURCES += \
    file.qrc

RC_FILE = icon.rc
