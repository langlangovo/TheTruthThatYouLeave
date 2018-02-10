#ifndef PIECEWORKDATA_H
#define PIECEWORKDATA_H
#include <QStringList>

struct PieceworkData
{
    QStringList name;       //人名
    QStringList number;     //手持终端编号
    QStringList trackingNO; //货运单号
};

struct RegisteredPerson
{
    QStringList number;     //手持终端编号
    QStringList name;       //人名
};

#endif // PIECEWORKDATA_H
