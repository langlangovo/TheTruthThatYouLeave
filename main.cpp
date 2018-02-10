#include "mainwindow.h"
#include <QApplication>
#include <QPixmap>
#include <QSplashScreen>
#include <QDateTime>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    QPixmap pixmap(":/images/Logo III.png");
    QSplashScreen splash(pixmap);
    splash.show();
    a.processEvents();
    Qt::Alignment topRight = Qt::AlignRight|Qt::AlignTop;
    splash.showMessage("正在启动..",topRight,Qt::white);
    MainWindow w;
    w.show();

    splash.finish(&w);

    return a.exec();
}
