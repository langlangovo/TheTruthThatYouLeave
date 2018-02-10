#include "excelthread.h"
#include "excel.h"
#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QDebug>
#include <QRegExp>
#include <QMessageBox>
#include <QProcess>
#include "pieceworkdata.h"
#include "windows.h"


ExcelThread::ExcelThread()
{
    debug = false;
}

void ExcelThread::setDebug(const bool state)
{
    debug = state;
}

void ExcelThread::setTxtFilePath(const QString &path)
{
    sourceDirName = path;
}

void ExcelThread::setTemplatePath(const QString &path)
{
    templatePath = path;
}

void ExcelThread::setOutFileName(const QString &path)
{
    outFileName = path;
}

void ExcelThread::run()
{
    qDebug()<<"sourceDirName"<<sourceDirName;
    qDebug()<<"templatePath"<<templatePath;
    qDebug()<<"outFileName"<<outFileName;

    QVariant var;
    bool success;
    success = input(var);
    if(!success)
    {
        emit returnResult("input: 读取错误");
        return;
    }

    success = output(var,outFileName);
    if(!success)
    {
        emit returnResult("output: 输出Excel文件错误");
        return;
    }

    if (success)
        emit returnResult("转换完成！");
    else
        emit returnResult("转换失败！");
}

bool ExcelThread::output(const QVariant &var, QString outFileName)
{
    emit returnResult("初始化Excel控件");
    Excel excel;
    bool success;
    emit returnResult("打开Excel模板:"+templatePath);
    success = excel.open(templatePath);

    if(!success){
        emit returnResult("链接Excel控件失败");
        Sleep(1000);
        return false;
    }
    if(debug)
        excel.display(true);
    emit returnResult("设置数据源");
    excel.setValue(var);
    emit returnResult("存储文件:"+outFileName);
    excel.saveAs(outFileName);
    emit returnResult("释放Excel控件");
    excel.release();
    return true;
}

bool ExcelThread::geTxtFileNameList(QStringList &txtFileNameList)
{
    QDir dir;
    QRegExp rx("\\S+.txt");
    dir.setPath(sourceDirName);
    QStringList fileList=dir.entryList();
    for(int i=0; i<fileList.size(); i++)
    {
        rx.indexIn(fileList.at(i));
        if(!rx.cap(0).isEmpty())
        {
            QString fileName = sourceDirName + "\\" + rx.cap(0);
            fileName.replace(QRegExp("/"), "\\");
            txtFileNameList.append(fileName);
        }
    }
    return true;
}

bool ExcelThread::readTxt(const QStringList txtFileNameList, PieceworkData *data)
{
    //读取所有Txt文本
    QRegExp rx_trackingNO("(\\d{12,13})");
    /*
    QRegExp rx_repace(
                "(\\d{12,13})\\t"
                "\\d\\d\\d\\d-\\d\\d-\\d\\d "
                "\\d\\d:\\d\\d:\\d\\d\t");
    */
    QRegExp rx_repace(
                "(\\d{12,13})\\t");
    for(int i=0; i<txtFileNameList.size(); i++)
    {
        QFile file(txtFileNameList.at(i));
        if ( file.open( QIODevice::ReadOnly ) ) {
            qDebug()<<"正在处理 "<<txtFileNameList.at(i)<<endl;
            QTextStream stream(&file);
            QString line;

            while ( !stream.atEnd() ) {
                line = stream.readLine();           // 不包括“\n”的一行文本
                rx_trackingNO.indexIn(line);

                data->trackingNO.append(rx_trackingNO.cap(0));
                rx_repace.indexIn(line);
                data->number.append(line.replace(rx_repace,""));
            }
            file.close();
        }else
            return false;

    }
    return true;
}

bool ExcelThread::swap(PieceworkData *pieData,QVariant &var)
{
    /*获取相应的编号*/
    QStringList fileNameList;
    QDir dir;
    QRegExp rx("\\S+.txt");
    dir.setPath(sourceDirName);
    QStringList fileList=dir.entryList();
    for(int i=0; i<fileList.size(); i++)
    {
        rx.indexIn(fileList.at(i));
        if(!rx.cap(0).isEmpty())
        {
            QString fileName = rx.cap(0).replace(".txt","");
            fileNameList.append(fileName);
        }
    }

    /*提取*/
    emit returnResult("提取人员编号");
    RegisteredPerson regPerson;
    QRegExp rx_number("\\d{1,2}");
    for(int i=0;i<fileNameList.size();i++)
    {
        rx_number.indexIn(fileNameList.at(i));
        if(!rx_number.cap(0).isEmpty())
        {
            QString number = rx_number.cap(0);
            QString name = fileNameList.at(i);
            name = name.replace(rx_number,"");
            regPerson.number.append(number);
            regPerson.name.append(name+number);
            emit returnResult(regPerson.name.at(i)+"\t"+regPerson.number.at(i));
        }
    }

    /*更换编号*/
    for(int i=0; i<pieData->number.size(); i++)
    {
        for (int j=0;j<regPerson.number.size();j++)
        {
            if(pieData->number.at(i) == regPerson.number.at(j))
            {
                pieData->name.append(regPerson.name.at(j));
                break;
            }else if(j == regPerson.number.size()-1 && pieData->number.at(i) != regPerson.number.at(j))
            {
                QString error = QString("ERROR NUMBER: ")+pieData->number.at(i);
                pieData->name.append(error);
            }
        }

    }

    //将文本转换到QVariant
    QVariantList llist;
    for (int i=0; i<pieData->trackingNO.size(); i++)
    {
        QVariantList list;//切记，不能放在for循环外声明！否则处理到7000多条会崩溃！放在这里还对有利于效率提升2倍！！
        list<<pieData->trackingNO.at(i)<<pieData->name.at(i);
        llist.append(QVariant(list));
        if(debug)
        {
            qDebug()<<"第"<<i<<"条数据:"<<pieData->trackingNO.at(i)<<pieData->name.at(i);
        }
    }
    var = QVariant(llist);
}

/**
 * @brief ExcelThread::input
 * @param var
 * @return
 * 这个函数本来是只用于读取目录下所有文本文件，并
 * 转换到QVariant二维数组当中的，但内部集合了编号
 * 更换部分的代码，今后需要剔除这部分代码，单独一个
 * 函数来处理。并且这个函数的算法可以优化，去掉不需
 * 要的一些操作。
 */
bool ExcelThread::input(QVariant &var)
{
    QStringList txtFileNameList;
    PieceworkData pieData;


    this->geTxtFileNameList(txtFileNameList);
    this->readTxt(txtFileNameList,&pieData);
    this->swap(&pieData,var);

    return true;
}

