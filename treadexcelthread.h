#pragma execution_character_set("utf-8")
#ifndef TREADEXCELTHREAD_H
#define TREADEXCELTHREAD_H

#include <QObject>
#include <QThread>
#include <QStringList>
#include <qmap.h>
#include <QVariant>

class TReadExcelThread : public QThread
{
    Q_OBJECT
public:
    explicit TReadExcelThread(QObject *parent = nullptr);

public:
    QMap<QString, QStringList> failureCodeMap;
    QMap<QString, QList<QStringList>> importListMap;

private:
    void run();
    void readFailureCodeFile();
    QMap<QString, QList<QStringList>>  readImportFile();

signals:
    void readExcelFinished( QMap<QString, QStringList>* failureCodeMap, QMap<QString, QList<QStringList>>* importListMap);
    void finishedSignal();
};

#endif // TREADEXCELTHREAD_H
