#include "twriteexcelthread.h"

#include <QTextStream>
#include <QAxObject>
#include <QDebug>
#include <QList>
#include <QMap>
#include <QStringList>
#include <QException>

TWriteExcelThread::TWriteExcelThread(const QString& outputPath, const QList<QStringList>& allFailStrList, const QList<QStringList>& allPassStrList, const QMap<QString, QList<QStringList>>& importListMap, const QString& dateStr, const QString& projectName , const QMap<QString, QStringList> failureCodeMap, QObject *parent) : QThread(parent)
  , m_allPassStrList(allPassStrList), m_allFailStrList(allFailStrList)  , m_importListMap(importListMap), m_outputPath(outputPath), m_dateStr(dateStr), m_projectName(projectName), m_failureCodeMap(failureCodeMap)
{


}

void TWriteExcelThread::run()
{
    failureCodeKeyStrList = m_failureCodeMap.keys();
    orderList();
    writeSheetOne();
    writeSheetTwo();
    emit writedExcelFinished();
}

void TWriteExcelThread::orderList()
{
//    QList<QStringList> allStationPassList;
    for (int passStrListIndex = 0; passStrListIndex < m_allPassStrList.count(); passStrListIndex++ ) {
        QStringList allLogPassStrList = m_allPassStrList.at(passStrListIndex);

        qDebug() << allLogPassStrList.count();
        int allFailCount = 0;

        for (int i = 0; i < m_importListMap.count(); i++) {

            for (const QStringList& oneProjectPN : m_importListMap.first()) {
                if (oneProjectPN.at(0) == allLogPassStrList.at(0)) {
                    allLogPassStrList.append(oneProjectPN.at(1));
                    goto end;
                }
            }
        }
        end:

        for (const QStringList& allLogFailStrList: m_allFailStrList) {
            QString serialIdStr = allLogPassStrList.at(0);
            if (allLogFailStrList.at(0) == serialIdStr ) {
                break;
            }
            allFailCount++;
        }

        if (allFailCount == m_allFailStrList.count() ) {
            allStationPassList.append(allLogPassStrList);
        }

    }


//    QList<QStringList> allStationFailList;
    for (int FailStrListIndex = 0; FailStrListIndex < m_allFailStrList.count(); FailStrListIndex++ ) {
        QStringList allLogFailStrList = m_allFailStrList.at(FailStrListIndex);
//    for (const QStringList& allLogFailStrList : m_allFailStrList) {

        int allPassCount = 0;
        for (const QStringList& allStationPassStrList : allStationPassList) {

            QString serialIdStr = allLogFailStrList.at(0);
            if (allStationPassStrList.at(0) == serialIdStr) {
                break;
            }
            allPassCount++;
        }
        if (allPassCount == allStationPassList.count() ) {

            bool stopLoop = false;
            for (int i = 0; i < m_importListMap.count(); i++) {

                for (const QStringList& oneProjectPN : m_importListMap.first()) {
                    if (oneProjectPN.at(0) == allLogFailStrList.at(0)) {
                        allLogFailStrList.append(oneProjectPN.at(1));
//                        goto end2;
                        stopLoop = true;
                    }
                }
                if (stopLoop) {
                    break;
                }
            }

            allStationFailList.append(allLogFailStrList);
//            end2:
        }

    }
}

/**
 * @brief TWriteExcelThread::编写Sheet 1
 */
void TWriteExcelThread::writeSheetOne()
{

    try {
        // 创建Excel应用程序对象
        QAxObject excel("Excel.Application");
        excel.setProperty("Visible", false); // 可以设置为true以显示Excel界面

        // 获取工作簿集合
        QAxObject *workbooks = excel.querySubObject("Workbooks");

        if (!workbooks) {
            qDebug() << "Failed to get workbooks.";
            emit writedExcelError("Failed to get workbooks.");

        }


        // 打开现有的Excel文件
        QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", QVariant(m_outputPath));
        if (!workbook) {
            qDebug() << "Failed to open workbook.";
            emit writedExcelError("Failed to get workbook.");
        }

        // 获取工作表集合
        QAxObject *worksheets = workbook->querySubObject("WorkSheets");
        if (!worksheets) {
            qDebug() << "Failed to get worksheets.";
            emit writedExcelError("Failed to get worksheets.");
        }

        // 获取第一个工作表
        QAxObject *worksheet = worksheets->querySubObject("Item(int)", 1);
        if (!worksheet) {
            qDebug() << "Failed to get worksheet.";
            emit writedExcelError("Failed to get worksheet.");
        }

        int row_index = 3;
        for (const QStringList& failStrList : allStationFailList) {
            QString snNum = failStrList.at(0);
            QString pnNum = ""; //  failStrList.last();

            if (snNum == "EMPTY") {
                continue;
            }

            if (failStrList.size() >= 5) {

                pnNum = failStrList.last();

            }

            // 在指定的单元格中填写内容
            QAxObject *cellIn1 = worksheet->querySubObject("Cells(int,int)", row_index, 1);     // 在row_index行,1列表格中填写内容
            if (cellIn1) {
                cellIn1->dynamicCall("SetValue(const QVariant&)", QVariant(m_dateStr));

            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell.");
            }

            QAxObject *cellIn2 = worksheet->querySubObject("Cells(int,int)", row_index, 2);
            if (cellIn2) {
                cellIn2->dynamicCall("SetValue(const QVariant&)", QVariant(m_projectName));

            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell 2.");
            }

            QAxObject *cellIn3 = worksheet->querySubObject("Cells(int,int)", row_index, 3);
            if (cellIn3) {
                cellIn3->dynamicCall("SetValue(const QVariant&)", QVariant(pnNum));

            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell 3.");
            }

            QAxObject *cell = worksheet->querySubObject("Cells(int,int)", row_index, 4);
            if (cell) {
                cell->dynamicCall("SetValue(const QVariant&)", QVariant(snNum));

            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell 4.");
            }

            QAxObject *cellResult = worksheet->querySubObject("Cells(int,int)", row_index, 5);
            if (cellResult) {
                cellResult->dynamicCall("SetValue(const QVariant&)", QVariant("fail"));
            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell 5.");
            }

            row_index++;
        }


        for (const QStringList& passStrList : allStationPassList) {
            QString snNum = passStrList.at(0);
            QString pnNum = passStrList.last();

            if (snNum == "EMPTY") {
                continue;
            }

            // 在指定的单元格中填写内容
            QAxObject *cellIn1 = worksheet->querySubObject("Cells(int,int)", row_index, 1);
            if (cellIn1) {
                cellIn1->dynamicCall("SetValue(const QVariant&)", QVariant(m_dateStr));

            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell pass 1.");
            }


            QAxObject *cellIn2 = worksheet->querySubObject("Cells(int,int)", row_index, 2);
            if (cellIn2) {
                cellIn2->dynamicCall("SetValue(const QVariant&)", QVariant(m_projectName));

            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell pass 2.");
            }

            QAxObject *cellIn3 = worksheet->querySubObject("Cells(int,int)", row_index, 3);
            if (cellIn3) {
                cellIn3->dynamicCall("SetValue(const QVariant&)", QVariant(pnNum));

            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell pass 3.");
            }

            QAxObject *cell = worksheet->querySubObject("Cells(int,int)", row_index, 4);
            if (cell) {
                cell->dynamicCall("SetValue(const QVariant&)", QVariant(snNum));

            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell pass 4.");
            }

            QAxObject *cellResult = worksheet->querySubObject("Cells(int,int)", row_index, 5);
            if (cellResult) {
                cellResult->dynamicCall("SetValue(const QVariant&)", QVariant("pass"));
            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Failed to get cell pass 5.");
            }

            row_index++;
        }

        QAxObject *dateCell = worksheet->querySubObject("Cells(int,int)", 3, 7);
        if (dateCell) {
            dateCell->dynamicCall("SetValue(const QVariant&)", QVariant(m_dateStr));

        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Failed to get cell pass 7.");
        }

        QAxObject *pnCell = worksheet->querySubObject("Cells(int,int)", 3, 8);
        if (pnCell) {
            pnCell->dynamicCall("SetValue(const QVariant&)", QVariant(m_projectName));

        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Failed to get cell pass 8.");
        }

        QAxObject *inputDescCell = worksheet->querySubObject("Cells(int,int)", 3, 9);
        if (inputDescCell) {
            int result = allStationPassList.count() + allStationFailList.count();
            inputDescCell->dynamicCall("SetValue(const QVariant&)", QVariant(result));

        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Failed to get cell pass 9.");
        }

        QAxObject *passDescCell = worksheet->querySubObject("Cells(int,int)", 3, 10);
        if (passDescCell) {
            passDescCell->dynamicCall("SetValue(const QVariant&)", QVariant(allStationPassList.count()));

        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Failed to get cell pass 10.");
        }

        QAxObject *failDescCell = worksheet->querySubObject("Cells(int,int)", 3, 11);
        if (failDescCell) {
            failDescCell->dynamicCall("SetValue(const QVariant&)", QVariant(allStationFailList.count()));

        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Failed to get cell pass 11.");
        }

        // 保存并关闭工作簿（可选）
        workbook->dynamicCall("SaveAs(const QString&)", m_outputPath);
        workbook->dynamicCall("Close()");
    }  catch (const std::exception& e) {
        // 捕获到std::exception类型的异常
            qDebug() << "Caught exception: " << e.what();
            emit writedExcelError(e.what());
    }

}

/**
 * @brief TWriteExcelThread::编写Sheet 2
 */
void TWriteExcelThread::writeSheetTwo()
{
    // 创建Excel应用程序对象
    QAxObject excel("Excel.Application");
    excel.setProperty("Visible", false); // 可以设置为true以显示Excel界面

    // 获取工作簿集合
    QAxObject *workbooks = excel.querySubObject("WorkBooks");
    if (!workbooks) {
        qDebug() << "Failed to get workbooks.";
        emit writedExcelError("Sheet two Failed to get workbooks.");
    }


    // 打开现有的Excel文件
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", QVariant(m_outputPath));
    if (!workbook) {
        qDebug() << "Failed to open workbook.";
        emit writedExcelError("Sheet two Failed to get workbook.");
    }

    // 获取工作表集合
    QAxObject *worksheets = workbook->querySubObject("WorkSheets");
    if (!worksheets) {
        qDebug() << "Failed to get worksheets.";
        emit writedExcelError("Sheet two Failed to get worksheets.");
    }

    // 获取第一个工作表
    QAxObject *worksheet = worksheets->querySubObject("Item(int)", 2);
    if (!worksheet) {
        qDebug() << "Failed to get worksheet.";
        emit writedExcelError("Sheet two Failed to get worksheet.");
    }


    int row_index = 3;
    for (const QStringList& failStrList : allStationFailList) {
        QString snNum = failStrList.at(0);
        QString pnNum = "";
        QString stationStr = failStrList.at(1);
        QString errorCodeStr =failStrList.last();
//        QString errorCodeStr = failStrList.at(1);

        if (snNum == "EMPTY") {
            continue;
        }

        if (failStrList.size() >= 5) {
            errorCodeStr = failStrList.at(failStrList.size() - 2);
            pnNum = failStrList.last();

        }


        // 在指定的单元格中填写内容
        QAxObject *cellIn1 = worksheet->querySubObject("Cells(int,int)", row_index, 1);     // 在row_index行,1列表格中填写内容
        if (cellIn1) {
            cellIn1->dynamicCall("SetValue(const QVariant&)", QVariant(m_dateStr));

        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Sheet two Failed to get worksheet.");
        }

        QAxObject *cellIn2 = worksheet->querySubObject("Cells(int,int)", row_index, 2);
        if (cellIn2) {
            cellIn2->dynamicCall("SetValue(const QVariant&)", QVariant(m_projectName));

        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Sheet two Failed to get cell 2.");
        }

        QAxObject *cellIn3 = worksheet->querySubObject("Cells(int,int)", row_index, 3);
        if (cellIn3) {
            cellIn3->dynamicCall("SetValue(const QVariant&)", QVariant(pnNum));

        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Sheet two Failed to get cell 3.");
        }

        QAxObject *cell = worksheet->querySubObject("Cells(int,int)", row_index, 4);
        if (cell) {
            cell->dynamicCall("SetValue(const QVariant&)", QVariant(snNum));

        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Sheet two Failed to get cell 4.");
        }

        QAxObject *cellResult = worksheet->querySubObject("Cells(int,int)", row_index, 5);
        if (cellResult) {
//             QString desc = failStrList.at(1);  //  不良工站
            cellResult->dynamicCall("SetValue(const QVariant&)", QVariant(stationStr));
        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Sheet two Failed to get cell 5.");
        }

        QAxObject *cellResult2 = worksheet->querySubObject("Cells(int,int)", row_index, 6);
        if (cellResult2) {
             QString errorCode = errorCodeStr; //  不良代码
            cellResult2->dynamicCall("SetValue(const QVariant&)", QVariant(errorCode));
        } else {
            qDebug() << "Failed to get cell.";
            emit writedExcelError("Sheet two Failed to get cell 6.");
        }

        if (failStrList.size() == 9) {



            QAxObject *cellResult3 = worksheet->querySubObject("Cells(int,int)", row_index, 7);
            if (cellResult3) {
                 QString errorCode = failStrList.at(4); //  不良项
                cellResult3->dynamicCall("SetValue(const QVariant&)", QVariant(errorCode));
            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Sheet two Failed to get cell 7.");
            }
            QAxObject *cellResult4 = worksheet->querySubObject("Cells(int,int)", row_index, 8);
            if (cellResult4) {
                 QString errorDesc = failStrList.at(5); //  不良描述
                cellResult4->dynamicCall("SetValue(const QVariant&)", QVariant(errorDesc));
            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Sheet two Failed to get cell 8.");
            }
            QAxObject *cellResult5 = worksheet->querySubObject("Cells(int,int)", row_index, 10);
            if (cellResult5) {
                 QString errorDesc = failStrList.at(6); //  维修动作
                cellResult5->dynamicCall("SetValue(const QVariant&)", QVariant(errorDesc));
            } else {
                qDebug() << "Failed to get cell.";
                emit writedExcelError("Sheet two Failed to get cell 10.");
            }
        }


//        if (failStrList.size() >= 6) {
//            QAxObject *cellResult = worksheet->querySubObject("Cells(int,int)", row_index, 5);
//            if (cellResult) {
//                 QString desc = failStrList.at(4);
//                cellResult->dynamicCall("SetValue(const QVariant&)", QVariant(desc));
//            } else {
//                qDebug() << "Failed to get cell.";
//            }

//            QAxObject *cellResult6 = worksheet->querySubObject("Cells(int,int)", row_index, 6);
//            if (cellResult6) {
//                QString actionDesc = failStrList.at(5);
//               cellResult6->dynamicCall("SetValue(const QVariant&)", QVariant(actionDesc));
//            } else {
//                qDebug() << "Failed to get cell.";
//            }
//        }


        row_index++;
    }




    // 保存并关闭工作簿（可选）
    workbook->dynamicCall("SaveAs(const QString&)", m_outputPath);
    workbook->dynamicCall("Close()");
}


