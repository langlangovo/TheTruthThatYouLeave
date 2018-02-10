#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "excel.h"
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QFileDialog>
#include <QDebug>
#include <QRegExp>
#include <QMessageBox>
#include <QProcess>
#include <QDateTime>
#include <QCloseEvent>
#include <windows.h>
#include <QMediaPlayer>

bool debug = false;
bool txtSwapExcel(QString sourceFileName,QString outFileName);

void MainWindow::version()
{
    QMessageBox::about(this,
                       "版本",
                       "2017-04-23  更新内容                          \n"
                       "——————————————————————————————————————————  \n"
                       "1.增加多线程的支持                             \n"
                       "\n\n"
                       "2017-04-26 更新内容"
                       "——————————————————————————————————————————  \n"
                       "1.添加通知“转换完成”&“转换失败“                  \n"
                       "2.支持多个文件同时操作，并输出到一个“输出.xls”文件  \n"
                       "\n\n"
                       "2017-04-29 更新内容"
                       "——————————————————————————————————————————  \n"
                       "1.优化转换速度，处理10000条数据，只需2-3秒        \n"
                       "2.输出文件名以当【前时间】命名                    \n"
                       "\n\n"
                       "2017-07-04 更新内容                           \n"
                       "——————————————————————————————————————————  \n"
                       "1.修复重复开多线程                              \n"
                       "\n\n"
                       "2017-10-28 更新内容                           \n"
                       "——————————————————————————————————————————  \n"
                       "1.增加人员编号自动识别功能\n"
                       "2.修复已知bug\n\n"
                       );
}

void MainWindow::about()
{
    QMessageBox::about(this,
                       "关于本软件",
                       "本软件的开发初衷是解放工作中的人们：               \n"
                       "    1.枯燥、繁琐的操作.                         \n"
                       "    2.误操作带来各种麻烦。                       \n"
                       "                                             \n"
                       "\n"
                       "——————————————————————————————————————————    \n"
                       "作者:langlangovo\n"
                       "QQ: 1300713226 \n"
                       "邮箱:langlangovo@163.com\n"
                       "[如果您想拥有一款类似的软件，来解决其它繁琐的工作，可以联系我定制，QQ请注明来意，谢谢！]"
                       );
}

void MainWindow::createMenuBar()
{
    QAction * aboutAction = new QAction(tr("关于"),this);
    ui->menuBar->addAction(aboutAction);
    connect(aboutAction,SIGNAL(triggered(bool)),this,SLOT(about()));

    QAction *versionAction = new QAction(tr("版本"),this);
    ui->menuBar->addAction(versionAction);
    connect(versionAction,SIGNAL(triggered(bool)),this,SLOT(version()));
}

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    QDir dir;
    QString musicPath = dir.currentPath() + "/EXIT TRANCE.mp3";
    qDebug()<<musicPath;
    QMediaPlayer *player;
    player = new QMediaPlayer;
    connect(player, SIGNAL(positionChanged(qint64)), this, SLOT(positionChanged(qint64)));
    player->setMedia(QUrl::fromLocalFile(musicPath));
    player->setVolume(50);
    player->play();

    Sleep(3000);
    ui->setupUi(this);
    setFixedSize(this->width(),this->height());
    setWindowTitle("The truth that you leave.");
    ui->statusBar->showMessage("langlangovo@163.com");
    createMenuBar();
    connect(&excelThread,SIGNAL(returnResult(QString)),
            this,SLOT(statusBarShowMessage(QString)));
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pushButton_selectDir_clicked()
{
    ui->pushButton_selectDir->setEnabled(false);
    //检查设置模板绝对路径
    QDir dir;
    templatePath = dir.currentPath() + "/模板.xls";
    QFileInfo file(templatePath);
    if(!file.isFile())
    {
        QMessageBox::warning(
                    this,
                    "警告！",
                    "模板不存在！\n\t"+templatePath,
                    QMessageBox::Ok);
    }else
    {
        excelThread.setTemplatePath(templatePath);
    }

    //设置要转换的Txt目录
    sourceDirName = QFileDialog::getExistingDirectory(
                this,
                "选择要转换的数据目录",
                "");
    if(sourceDirName.isEmpty())
    {
        ui->statusBar->showMessage("已取消选择");
        ui->pushButton_selectDir->setEnabled(true);
        return;
    }
    ui->statusBar->showMessage("正在转换数据，请稍等片刻...");
    excelThread.setTxtFilePath(sourceDirName);

    //设置输出文件名
    QDateTime nowTime = QDateTime::currentDateTime();
    QString outFileName = sourceDirName + "/" + nowTime.toString("yyyy-MM-dd ddd hhmmss") + ".xls";
    excelThread.setOutFileName(outFileName);
    excelThread.start();

}

void MainWindow::on_pushButton_clicked()
{
   static bool checked = false;

   if(checked == false){
       ui->pushButton->setText("关闭Debug");
       debug = true;
       excelThread.setDebug(true);
       checked = true;

   }else
   {
       ui->pushButton->setText("开启Debug");
       debug = false;
       excelThread.setDebug(false);
       checked = false;
   }
}

void MainWindow::closeEvent(QCloseEvent *event)
{
    if(excelThread.isRunning()){
        int ret = QMessageBox::warning(
                    this,
                    "警告！",
                    "数据尚未转换完成，是否强制退出？",
                    QMessageBox::Yes|QMessageBox::No);
        if(ret == QMessageBox::Yes)
        {
            excelThread.quit();
            excelThread.wait();
            event->accept();
        }
        else
            event->ignore();
    }
}

 void MainWindow::statusBarShowMessage(QString message)
 {
     ui->textEdit_debug->append(message);
     ui->statusBar->showMessage(message);
     /*
     QMessageBox::information(
                 this,
                 "提示",
                 message,
                 QMessageBox::Ok);
     */
     if (message == "转换完成！" | message == "转换失败！")
         ui->pushButton_selectDir->setEnabled(true);

 }

 bool MainWindow::loadSettings()
 {
     //检查设置模板绝对路径
     QDir dir;
     templatePath = dir.currentPath() + "/模板.xls";
     QFileInfo file(templatePath);
     if(!file.isFile())
     {
         QMessageBox::warning(
                     this,
                     "警告！",
                     "模板不存在！\n\t"+templatePath,
                     QMessageBox::Ok);
         templatePath = "";
         return false;
     }

     QString settingsFilePath = dir.currentPath() + QString("/settings.ini");
     QFile settingsFile(settingsFilePath);
     if (!settingsFile.open(QIODevice::ReadOnly | QIODevice::Text))
     {
         ui->statusBar->showMessage("配置文件不存在！");
         return false;
     }
     QTextStream in(&settingsFile);
     if (!in.atEnd()) {
         QString line = in.readLine();
         sourceDirName = line;
         qDebug()<<line;
     }else
     {
         ui->statusBar->showMessage("配置文件是空的.");
         return false;
     }
     settingsFile.close();


     //设置输出文件名
     QDateTime nowTime = QDateTime::currentDateTime();
     QString outFileName = sourceDirName + "/" + nowTime.toString("yyyy-MM-dd ddd hhmmss") + ".xls";

     excelThread.setTemplatePath(templatePath);
     excelThread.setTxtFilePath(sourceDirName);
     excelThread.setOutFileName(outFileName);
     qDebug()<<templatePath;
     qDebug()<<sourceDirName;
     qDebug()<<outFileName;

     return true;
 }

