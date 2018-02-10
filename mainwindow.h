#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "excelthread.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_pushButton_selectDir_clicked();
    void on_pushButton_clicked();
    void about();
    void version();
    void createMenuBar();
    void closeEvent(QCloseEvent *event);
    void statusBarShowMessage(QString message);

private:
    Ui::MainWindow *ui;
    ExcelThread excelThread;
    QString templatePath;       //模板路径
    QString sourceDirName;      //Txt文件源路径
    QString outFileName;        //输出文件名

    bool loadSettings();        //加载设置
    bool createSettings();      //创建配置文件
};

#endif // MAINWINDOW_H
