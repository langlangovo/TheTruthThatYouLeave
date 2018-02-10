#ifndef EXCEL_H
#define EXCEL_H
#include <QString>
#include <QAxObject>
#include <QVariantList>

/**
 ******************************************************************************
 * @brief The Excel class
 * 项目要添加：
 *1.
 *  QT       += axcontainer
 *
 *2.
 *  a.setValue 效率要比 setCellValue 要高至少10倍
 *  b.setValue 是一个二维数组，原形是QVariant<QVariantList<QVariant<QVariantList>>>
 *
 * 3.
 *   操作完Excel，一定要记得释放Excel！否则输出文件会挂起到“只读状态”，并且Excel进程一直挂起
 *   消耗系统内存。每启动一次，就多挂起一个Excel
 * ****************************************************************************
 * langlangovo 2017-04-29
 */

class Excel
{
public:
    Excel();                                                        //连接Excel控件
    ~Excel();
    bool open(const QString fileName);                              //打开Excel文件
    bool create(const QString fileName);                            //创建Excel文件
    bool appendSheet(const QString sheetName);                      //添加1个worksheet
    bool setCellValue(int row, int column, const QString &value);   //向单元表格写入数据
    QVariant readAll();                                             //读取所有数据
    bool saveAs(QString fileName);                                  //另存文件
    bool save();                                                    //保存文件
    bool display(bool state);                                       //显示Excel 窗口
    bool setValue(const QVariant var);                              //二维数组输出
    bool swapWin32FilePath(QString &filePath);                      //转换成windows的路径风格
    void release();                                                 //释放Excel组件
 private:
    QAxObject *pApplication = NULL;
    QAxObject *pWorkBooks   = NULL;
    QAxObject *pWorkBook    = NULL;
    QAxObject *pWorkSheets  = NULL;
    QAxObject *pWorkSheet   = NULL;

    QString fileName;                                               //open create函数的文件路径
    QString to26AlphabetString(int row, int column);                //数字转换成字母
};

#endif // EXCEL_H
