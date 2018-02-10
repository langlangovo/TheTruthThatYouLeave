#include "excel.h"
#include <QFile>
#include <QDir>
#include <QList>
#include <QVariant>
#include <QDebug>

Excel::Excel()
{
}

/**
 * @brief Excel::create
 * @param fileName 一定要是绝对路径
 * @return
 */
bool Excel::create(const QString fileName)
{
    bool success;
    Excel::fileName = fileName;
    swapWin32FilePath(Excel::fileName);

    release();                                              //先释放再创建的
    pApplication = new QAxObject;
    success = pApplication->setControl("Excel.Application");//连接Excel控件
    if(!success)
        return false;
    pApplication->setProperty("Visible", false);

    pApplication->setProperty("DisplayAlerts", false);		//不显示任何警告信息。
    pWorkBooks = pApplication->querySubObject("Workbooks"); //获取工作簿集合
    pWorkBooks->dynamicCall("Add");                         //添加WorkBook
    pWorkBook = pApplication->querySubObject(("ActiveWorkBook"));
    pWorkSheets = pWorkBook->querySubObject("Sheets");
    pWorkSheet = pWorkSheets->querySubObject("Item(int)", 1);
    pWorkSheet->setProperty("Name", "langlangovo");             //设置标题

    return true;
}

/**
 * @brief Excel::open
 * @param fileName 一定要是绝对路径
 * @return
 */
bool Excel::open(const QString fileName)
{
    Excel::fileName = fileName;
    bool success;
    swapWin32FilePath(Excel::fileName);
    QFile file(fileName);
    if(!file.exists())
        return false;

    release();
    pApplication = new QAxObject;
    success = pApplication->setControl("Excel.Application");        //连接Excel控件
    if(!success)
        return false;
    pApplication->setProperty("Visible", true);                    //显示窗口
    pApplication->setProperty("DisplayAlerts", true);              //不显示任何警告信息。

    pWorkBooks = pApplication->querySubObject("Workbooks");         //获取工作簿集合
    pWorkBooks->dynamicCall("Open(const QString &)", fileName);
    pWorkBook = pApplication->querySubObject("ActiveWorkBook");
    pWorkSheets = pWorkBook->querySubObject("Sheets");
    pWorkSheet = pWorkSheets->querySubObject("Item(int)", 1);
    return true;
}

/**
 * @brief Excel::save 保存文件
 * @return
 */
bool Excel::save()
{
    if(fileName.isEmpty())
        return false;
    pWorkBook->dynamicCall("SaveAs(const QString &)",fileName);
    return true;
}

/**
 * @brief Excel::saveAs
 * @param fileName 一定要是绝对路径
 * @return
 */
bool Excel::saveAs(QString fileName)
{
    swapWin32FilePath(fileName);
    pWorkBook->dynamicCall("SaveAs(const QString &)",fileName);

    return true;
}

/**
 * @brief Excel::release
 */
void Excel::release()
{
    if (pApplication != NULL)
      {
          pApplication->dynamicCall("Quit()");
          delete pApplication;
          pApplication = NULL;
      }
}

/**
 * @brief Excel::display
 * @param state bool 显示或不现实Excel窗口
 */
bool Excel::display(bool state)
{
    if(fileName.isEmpty())
        return false;
    pApplication->dynamicCall("SetVisible(bool)", state);
    return true;
}

Excel::~Excel()
{
    release();
}

/**
 * @brief Excel::setCellValue 设置单元表格值
 * @param row 表格行,不能<1
 * @param column 表格列,不能<1
 * @param value 坐标的值,不能为空
 * @return
 */
bool Excel::setCellValue(int row, int column, const QString &value)
{
    if (row<1 or column<1 )
        return false;

    if (value.isEmpty())
        return false;

    QAxObject *pRange = pWorkSheet->querySubObject("Cells(int,int)", row, column);
    pRange->setProperty("NumberFormat","@");
    pRange->dynamicCall("Value", value);
    return true;
}

/**
 * @brief Excel::appendSheet 添加一个工作表
 * @param sheetName 工作表名称,值不能为空
 * @return
 */
bool Excel::appendSheet(const QString sheetName)
{
    if (sheetName.isEmpty())
        return false;
    int cnt  = pWorkSheets->property("Count").toInt();
    QAxObject *pLastSheet = pWorkSheets->querySubObject("Item(int)", cnt);
    pWorkSheets->querySubObject("Add(QVariant)", pLastSheet->asVariant());
    pWorkSheet = pWorkSheets->querySubObject("Item(int)", cnt);
    pLastSheet->dynamicCall("Move(QVariant)", pWorkSheet->asVariant());
    pWorkSheet->setProperty("Name", sheetName);
    return true;
}

/**
 * @brief Excel::to26AlphabetString
 * @param row
 * @param column
 * @return
 * 这个函数只针对本软件，应该修改为只单纯的
 * 转换，需要修改！
 */
QString Excel::to26AlphabetString(int row, int column)
{
    column -= 1;
    row += 1;

    QString value;
    char alphabet = 'A';
    unsigned int i, ret, temp;

    //转换字母列
    i = column/26;
    temp = column%26;
    ret = alphabet + temp;
    for ( uint j=0; j<i; j++)
    {
        value += "A";
    }
    if (ret>0)
    {
        value += QString(ret);
    }

    //行号
    value += QString::number(row,10);

    //完整的Excel范围
    value = "A2:" + value;
    return value;

}

bool Excel::setValue(const QVariant var)
{
    QVariantList llist;
    QVariantList list;
    int row, column;

    //得到二维数组的大小
    llist = var.toList();
    row = llist.size();
    if (!row)
        return false;
    list = llist.at(0).toList();
    column = list.size();
    if (!column)
        return false;

    //通过二维数组大小，计算出Excel表格范围。
    qDebug()<<to26AlphabetString(row,column);
    //设置Excel 输出范围
    QAxObject *range = pWorkSheet ->querySubObject("Range(QString)", to26AlphabetString(row,column));
    //输出二维数组到Excel
    range->setProperty("Value", var);   //aline 是一个QList<QVaiant>变量
    return true;
}

bool Excel::swapWin32FilePath(QString &filePath)
{
    if(!filePath.isEmpty()){
        filePath.replace("/","\\");
        return true;
    }else
        return false;
}
