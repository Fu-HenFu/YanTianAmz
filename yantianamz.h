#pragma execution_character_set("utf-8")
#ifndef YANTIANAMZ_H
#define YANTIANAMZ_H

#include <QWidget>
#include <qdir.h>
#include <QLabel>
#include <qpushbutton.h>
#include <QMap>
#include "treadexcelthread.h"



QT_BEGIN_NAMESPACE
namespace Ui { class yantianamz; }
QT_END_NAMESPACE

class yantianamz : public QWidget
{
    Q_OBJECT

public:
    yantianamz(QWidget *parent = nullptr);
    ~yantianamz();

private:
    Ui::yantianamz *ui;
    QString outputFilePath;
    QMap<QString, QStringList> failureCodeMap;  //
    QMap<QString, QList<QStringList>> importListMap;
    QString logDatetimeStr;
    QString projectName;
    TReadExcelThread* readThread;
    QWidget* converWidget;
    QPushButton* selectBtn;
    QLabel* tipsLabel;

public slots:
    void selectFileSlot();
    void readFinishedSlot();
    void readExcelFinishedSlot(QMap<QString, QStringList>* failureCodeMap, QMap<QString, QList<QStringList>>* importListMap);

private:
    bool compareStringLists(const QStringList& list1, const QStringList& list2);
    void readFailureCodeFile();
    QMap<QString, QList<QStringList>> readImportFile();
    void findFilesWithKeyword(const QDir& dir, const QString& keyword, QStringList& filePaths);
    QMap<QString, QList<QStringList>> readCsvFile(const QString& filePath);
    int writeFile( QList<QStringList> allFailStrList,  QList<QStringList> allPassStrList);
    void writeExcel(const QList<QStringList>& allFailList);
};
#endif // YANTIANAMZ_H
