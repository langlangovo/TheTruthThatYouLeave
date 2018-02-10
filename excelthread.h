#ifndef EXCELTHREAD_H
#define EXCELTHREAD_H
#include <QThread>
#include "excel.h"
#include "pieceworkdata.h"

class ExcelThread: public QThread
{
    Q_OBJECT
public:
    ExcelThread();
    void setTxtFilePath(const QString &path);             //设置文本文件夹绝对路径
    void setTemplatePath(const QString &path);            //设置转换模板绝对路径
    void setOutFileName(const QString &path);             //设置输出文件绝对路径
    void setDebug(const bool state);                      //设置Debug模式

protected:
    void run();

signals:
    void returnResult(QString );

private:
    QString sourceDirName;                                //文本文件夹绝对路径
    QString templatePath;                                 //输出文件的模板绝对路径
    QString outFileName;                                  //输出文件路径名
    volatile bool debug;                                  //开关debug模式                               

    bool input(QVariant &var);                                         //读取所有文本pieceworkData结构体
    bool output(const QVariant &var, QString outFileName);//输出Excel
    bool swap(PieceworkData *pieData, QVariant &var);   //转换数据到var数组

    bool geTxtFileNameList(QStringList &txtFileNameList); //获取txt文本文件列表
    bool readTxt(const QStringList txtFileNameList, PieceworkData *data);//读取txt数据
};

#endif // EXCELTHREAD_H
