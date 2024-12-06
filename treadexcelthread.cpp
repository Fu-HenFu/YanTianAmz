#include "treadexcelthread.h"

#include <QStandardPaths>
#include <QDebug>
#include <QAxObject>
#include <QCoreApplication>
//#include "xlsxdocument.h"
#include <QDesktopServices>
#include <QTextCodec>
#include <windows.h> // For CoInitialize and CoUninitialize

#pragma execution_character_set("utf-8")
//using namespace QXlsx;

TReadExcelThread::TReadExcelThread(QObject *parent) : QThread(parent)
{


}

void TReadExcelThread::run()
{
    readFailureCodeFile();
    importListMap = readImportFile();

    emit readExcelFinished(&failureCodeMap, &importListMap);
    emit finishedSignal();

}


/**
 * @brief yantiancnd::readFailureCodeFile 读取记录了错误码的文件.取出值,并保留
 */
void TReadExcelThread::readFailureCodeFile()
{

    // 初始化 COM 库
    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        qWarning() << "Failed to initialize COM library:" << hr;
        return ;
    }

    QString appInstalledPath = QCoreApplication::applicationDirPath();
    // 获取桌面文件夹的路径
    //    QString desktopPath = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation);
    appInstalledPath.append("/Rework_Failure_Codes_Ver.update_Needlefish.xlsx");


    // 创建Excel应用程序对象
    QAxObject excel("Excel.Application");
    excel.setProperty("Visible", false); // 可以设置为true以显示Excel界面

    // 获取工作簿集合
    QAxObject *workbooks = excel.querySubObject("WorkBooks");
    if (!workbooks) {
        qDebug() << "Failed to get workbooks.";
//        return -1;
    }


    // 打开现有的Excel文件
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", QVariant(appInstalledPath));
    if (!workbook) {
        qDebug() << "Failed to open workbook.";
//        return -1;
    }

    // 获取工作表集合
    QAxObject *worksheets = workbook->querySubObject("WorkSheets");
    if (!worksheets) {
        qDebug() << "Failed to get worksheets.";
//        return -1;
    }

    // 获取第一个工作表
    QAxObject *worksheet = worksheets->querySubObject("Item(int)", 1);
    if (!worksheet) {
        qDebug() << "Failed to get worksheet.";
//        return -1;
    }

    //获取该sheet的使用范围对象
    QAxObject * usedrange = worksheet->querySubObject("UsedRange");

    QAxObject * rows = usedrange->querySubObject("Rows");
    QAxObject * columns = usedrange->querySubObject("Columns");

    int nRow = rows->property("Count").toInt();
    int nCol = columns->property("Count").toInt();

//    qDebug() << "xls行数："<<nRow;
//    qDebug() << "xls列数："<<nCol;

    // 获取行数、列数
//    int nRow = worksheet->dynamicCall("UsedRange.Rows.Count").toInt();
//    int nCol = worksheet->dynamicCall("UsedRange.Columns.Count").toInt();

    // 读取单元格数据
    for (int i = 2; i <= nRow; i++) {
        QStringList failureStrList;
//            for (int j = 1; j <= nCol; j++) {
        QAxObject* cell = worksheet->querySubObject("Cells(int,int)", i, 3);
        QVariant cellValue = cell->property("Text");
        failureStrList.append(cellValue.toString());
        cellValue.toString(); // 输出单元格的值

        QAxObject* cell4 = worksheet->querySubObject("Cells(int,int)", i, 4);
        QVariant cell4Value = cell4->property("Text");
        failureStrList.append(cellValue.toString());

        QAxObject* cell5 = worksheet->querySubObject("Cells(int,int)", i, 5);
        QVariant cell5Value = cell5->property("Text");
        failureStrList.append(cell5Value.toString());

        QAxObject* cell6 = worksheet->querySubObject("Cells(int,int)", i, 6);
        QVariant cell6Value = cell6->property("Text");
        failureStrList.append(cell6Value.toString());

        QAxObject* cell8 = worksheet->querySubObject("Cells(int,int)", i, 8);
        QVariant cell8Value = cell8->property("Text");
        failureStrList.append(cell8Value.toString());
//            }

         failureCodeMap[cellValue.toString()] = failureStrList;
    }

    // 关闭工作簿和Excel对象
    workbook->dynamicCall("Close()");
    excel.dynamicCall("Quit()");

    // 反初始化 COM 库
    CoUninitialize();

}

/**
 * @brief yantianamz::readImportFile
 * @return 返回进口清单中,dsn对应的产品号
 */
QMap<QString, QList<QStringList> > TReadExcelThread::readImportFile()
{
     QMap<QString, QList<QStringList>> importMap;
    // 初始化 COM 库
    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        qWarning() << "Failed to initialize COM library:" << hr;
        return importMap;
    }

    QString appInstalledPath = QCoreApplication::applicationDirPath();
    // 获取桌面文件夹的路径
    //    QString desktopPath = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation);
    QString sheetName( tr("NF所有进口整机汇总.xlsx"));
    QTextCodec *code = QTextCodec::codecForName("GB2312");//解决中文路径问题
    QByteArray filePathByteArr = code->fromUnicode(sheetName);
    sheetName = code->toUnicode(filePathByteArr);
    appInstalledPath = appInstalledPath + "/" + sheetName;
    // 创建Excel应用程序对象
    QAxObject excel("Excel.Application", 0);
    excel.setProperty("Visible", false); // 可以设置为true以显示Excel界面

    // 获取工作簿集合
    QAxObject *workbooks = excel.querySubObject("WorkBooks");
    if (!workbooks) {
        qDebug() << "Failed to get workbooks.";
    }


    // 打开现有的Excel文件
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", QVariant(appInstalledPath));
    if (!workbook) {
        qDebug() << "Failed to open workbook.";
    }

    // 获取工作表集合
    QAxObject *worksheets = workbook->querySubObject("WorkSheets");
    if (!worksheets) {
        qDebug() << "Failed to get worksheets.";
//        return -1;
    }

    // 获取第一个工作表
    QAxObject *worksheet = worksheets->querySubObject("Item(int)", 1);
    if (!worksheet) {
        qDebug() << "Failed to get worksheet.";
//        return -1;
    }

    // 获取UsedRange对象
    QAxObject *usedRange = worksheet->querySubObject("UsedRange");

    // 获取行数
    QAxObject *rows = usedRange->querySubObject("Rows");
    int rowCount = rows->property("Count").toInt();

//        QMap<QString, QList<QStringList>> importMap;

    // 遍历每一行，检查隐藏状态，并读取数据
    for (int row = 2; row <= rowCount; ++row) {

        QStringList importStrList;
        QAxObject *rowObject = worksheet->querySubObject("Rows(int)", row);
        if (rowObject && !rowObject->property("Hidden").toBool()) {
            qDebug() << row << endl;
            QAxObject *cell = worksheet->querySubObject("Cells(int,int)", row, 2);
//                QVariant dd = cell->property("Value");
            importStrList.append(cell->property("Text").toString());
            QAxObject *cell3 = worksheet->querySubObject("Cells(int,int)", row, 3);
            QString partNoStr = cell3->property("Text").toString();
            importStrList.append(partNoStr);


            QList<QStringList> tempList = importMap[partNoStr];
            tempList.append(importStrList);
            importMap[partNoStr] = tempList;

        }
    }


    // 保存并关闭工作簿（可选）
    workbook->dynamicCall("SaveAs(const QString&)", sheetName);
    workbook->dynamicCall("Close()");

    // 反初始化 COM 库
    CoUninitialize();
    return importMap;
}


