#include "yantianamz.h"
#include "ui_yantianamz.h"
//#include "xlsxdocument.h"
#include <qpushbutton.h>
#include <qlayout.h>
#include <qdebug.h>
#include <qfile.h>
#include <QGraphicsOpacityEffect>

#include <QFileInfo>
#include <QFileInfoList>
#include <QFileDialog>
#include <QDesktopServices>
#include <QUrl>
#include <QTextStream>
#include <QAxObject>
#include <QDate>
#include <QStandardPaths>
#include <qstandardpaths.h>
#include <QTextCodec>
#include <QMessageBox>
#include <QRegularExpression>
#include "twriteexcelthread.h"
#include "treadexcelthread.h"

#pragma execution_character_set("utf-8")

//using namespace QXlsx;

yantianamz::yantianamz(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::yantianamz)
{
    ui->setupUi(this);
    QVBoxLayout* containerLayout = new QVBoxLayout(this);
    selectBtn = new QPushButton(this);
    connect(selectBtn, &QPushButton::clicked, this, &yantianamz::selectFileSlot);
    selectBtn->setText(tr("search dir"));

    tipsLabel = new QLabel(this);
    tipsLabel->setAlignment(Qt::AlignCenter);
    tipsLabel->setText(tr(""));

    containerLayout->addWidget(selectBtn);
    containerLayout->addStretch(1);
    containerLayout->addWidget(tipsLabel);

    containerLayout->addStretch(1);
    //QXlsx::Document xlsx;
    //xlsx.saveAs("Test3.xlsx");

    converWidget = new QWidget(this);
    converWidget->setFixedSize(this->size());
    QLabel* initailTips = new QLabel(converWidget);
    initailTips->setText(tr("Initially"));
    initailTips->setAlignment(Qt::AlignCenter);
    initailTips->setStyleSheet("background-color:transparent;") ;
    QVBoxLayout* converLayout = new QVBoxLayout();
    converLayout->addStretch(1);
    converLayout->addWidget(initailTips);
    converLayout->addStretch(1);
    converWidget->setLayout(converLayout);


    converWidget->setStyleSheet("background-color: rgba(0, 0, 0, 75);"); // 设置半透明黑色背景
    QGraphicsOpacityEffect *opacityEffect = new QGraphicsOpacityEffect;
    opacityEffect->setOpacity(0.3);
    converWidget->setGraphicsEffect(opacityEffect);
//    converWidget->show();



    readThread = new TReadExcelThread(this);
    QObject::connect(readThread, &TReadExcelThread::readExcelFinished, this, &yantianamz::readExcelFinishedSlot);
    QObject::connect(readThread, &TReadExcelThread::finishedSignal, this,&yantianamz::readFinishedSlot );
    readThread->start();
}

/**
 * @brief yantianamz::选择日志所在文件夹
 */
void yantianamz::selectFileSlot()
{
    // 获取桌面文件夹的路径
    QString desktopPath = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation);

    //QString folderPath = desktopPath; // 替换为你想要打开的文件夹路径
    // 设置对话框的标题
    QString dialogTitle = tr("chose dir");
    // 弹出对话框让用户选择文件夹
    QString folderPath = QFileDialog::getExistingDirectory(nullptr, dialogTitle, desktopPath);

    outputFilePath = folderPath;
    // 检查用户是否选择了文件夹
    if (!folderPath.isEmpty()) {
        qDebug() << "Selected folder:" << folderPath;
    }
    else {
        qDebug() << "No folder selected.";
    }
    // 使用 QUrl 从文件路径创建 URL
    //QUrl folderUrl = QUrl::fromLocalFile(folderPath);

    //// 使用 QDesktopServices 打开 URL（即文件夹）
    //bool success = QDesktopServices::openUrl(folderUrl);

    QStringList filePaths;
    QDir dirPath(folderPath);
    findFilesWithKeyword(dirPath, "amz", filePaths);

    QMap<QString,  QMap<QString, QList<QStringList>>> allCheckResultMap;

//    QMap<QString,  QList<QMap<QString, QList<QStringList>>>> allLogFailNPassMap;
    QList<QStringList> allLogPassList, allLogFailList;
//    QMap<QString,  QList<QMap<QString, QList<QStringList>>>> allLogPassMap;

//    QList<QStringList> allFailList;
    qDebug() << "FIle count" << filePaths.count() << endl;
    for (QString filePath : filePaths)
    {
        QFileInfo fileInfo(filePath);
        QString dirName = fileInfo.dir().dirName();
        QMap<QString, QList<QStringList>> oneLogFileList = readCsvFile(filePath);
        allCheckResultMap[dirName] = oneLogFileList;
//        QMap<QString, QList<QStringList>> tempFailMap, tempPassMap;
//        tempFailMap[dirName] = oneLogFileList["fail"];

        QList<QStringList> uniquePassListOfLists;
        for (int i = 0; i < oneLogFileList["pass"].size(); ++i) {
            bool isDuplicate = false;
            // 内层循环检查是否与之前的 QStringList 重复
            for (int j = 0; j < uniquePassListOfLists.size(); ++j) {
                if (compareStringLists(oneLogFileList["pass"][i], uniquePassListOfLists[j])) {
                    isDuplicate = true;
                    break;
                }
            }
            // 如果没有重复，则添加到结果列表中
            if (!isDuplicate) {
                uniquePassListOfLists.append(oneLogFileList["pass"][i]);
            }
        }

        allLogPassList.append(uniquePassListOfLists);
        allLogFailList.append(oneLogFileList["fail"]);

    }

    // 用于存储已经出现过的属性值
    QSet<QString> seenAttributes;
    // 用于记录重复的属性值及其重复次数
    QMap<QString, int> duplicateCounts;

    QList<QStringList> uniquePassStrList;
    // 遍历 QList<QStringList>
    for (const QStringList &logEntry : allLogPassList) {
        // 假设我们关心的属性是第一个字符串
        QString attribute = logEntry.first();

        // 检查这个属性值是否已经在 QSet 中
        if (seenAttributes.contains(attribute)) {
            // 如果在 QSet 中，说明有重复，增加重复计数
            duplicateCounts[attribute]++;
            qDebug() << "duplicateCounts有:" << duplicateCounts[attribute] << ":" << attribute << endl;
            if (duplicateCounts[attribute] == filePaths.count() - 1) {
                uniquePassStrList.append(logEntry);
            }
        } else {
            // 如果不在 QSet 中，将其添加到 QSet 中
            seenAttributes.insert(attribute);
        }
    }


    qDebug() << "总共有:" << uniquePassStrList.count() << "and" << allLogPassList.count() << endl;

    if (allCheckResultMap.count() > 0) {
        selectBtn->setEnabled(false);
        int outputResult = writeFile(allLogFailList, uniquePassStrList);
        tipsLabel->setText(tr("导出中..."));
        if (outputResult == -1) {

            selectBtn->setEnabled(true);
            QMessageBox* messageBox = new QMessageBox(this);
            messageBox->setWindowTitle(tr("中断提示"));
            messageBox->setText(tr("目录下没有导出文件模板"));
            messageBox->setIcon(QMessageBox::Warning);
            messageBox->addButton(QMessageBox::Ok);
            messageBox->exec();
        }
    }
}

void yantianamz::readFinishedSlot()
{
    qDebug() << "slot" << endl;
}

void yantianamz::readExcelFinishedSlot(QMap<QString, QStringList>* failureCodeMap, QMap<QString, QList<QStringList>>* importListMap)
{
    this->failureCodeMap = *failureCodeMap;
    this->importListMap = *importListMap;

    // 初始化完成，隐藏覆盖层
    converWidget->hide();
    converWidget->deleteLater();
}

/**
 * @brief yantianamz::比较两个QStringList内容是否相同
 * @param list1
 * @param list2
 * @return
 */
bool yantianamz::compareStringLists(const QStringList &list1, const QStringList &list2)
{
    // 辅助函数，用于比较两个 QStringList 是否相等

    if (list1.size() != list2.size()) {
        return false;
    }
//    for (int i = 0; i < list1.size(); ++i) {
        if (list1[0] != list2[0]) {
            return false;
        }
//    }
    return true;

}

/**
 * @brief 查找出文件名含offline的.csv文件.存入filePaths
 * @param dir
 * @param keyword
 * @param filePaths
 */
void yantianamz::findFilesWithKeyword(const QDir& dir, const QString& keyword, QStringList& filePaths)
{
    // 获取文件夹中的所有文件和子文件夹
    QFileInfoList fileList = dir.entryInfoList(QDir::Files | QDir::Dirs | QDir::NoDotAndDotDot, QDir::DirsFirst);

    // 遍历文件和文件夹
    for (const QFileInfo& fileInfo : fileList) {
        // 如果是文件夹，则递归查询
        if (fileInfo.isDir()) {
            findFilesWithKeyword(QDir(fileInfo.absoluteFilePath()), keyword, filePaths);
        }
        // 如果是文件，则检查文件名是否包含关键字
        else if (fileInfo.isFile() && fileInfo.fileName().contains(keyword, Qt::CaseInsensitive) && fileInfo.fileName().contains("offline", Qt::CaseInsensitive)) {

            // 如果包含关键字，则记录文件路径
            filePaths.append(fileInfo.absoluteFilePath());
            // 正则表达式匹配 yyyyMMDD 格式
            QRegularExpression re("(\\d{4})(\\d{2})(\\d{2})");
            QRegularExpressionMatch match = re.match( fileInfo.fileName());
            if (match.hasMatch()) {
            // 提取匹配的日期部分
                logDatetimeStr = match.captured(0);
                qDebug() << "字符串包含 yyyyMMDD 格式的日期" << logDatetimeStr;
            } else {
                qDebug() << "字符串不包含 yyyyMMDD 格式的日期";
            }
            QStringList fileNameSplitList = fileInfo.fileName().split("_");
            if (fileNameSplitList.count() > 3) {
                projectName = fileNameSplitList.at(2);
            }

        }
    }
}


/**
 * @brief yantiancnd::readCsvFile 读取csv文件
 * @param filePath
 * @return
 */
QMap<QString, QList<QStringList>> yantianamz::readCsvFile(const QString& filePath)
{
    QList<QString> failureCodeList = failureCodeMap.keys();
    QMap<QString, QList<QStringList>> passNFailList;
    QList<QStringList> failureList;
    QList<QStringList> passList;

    QFile file(filePath);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        qDebug() << "Error opening file for reading:" << filePath;
        return passNFailList;
    }

    bool currentRow = false;
    int test_result_column = -1;
    QTextStream in(&file);

    while (!in.atEnd()) {
        QStringList failStrList;
        QStringList passStrList;
        QString line = in.readLine();
        QStringList cellStrList = line.split(",");

        if (currentRow)
        {
            if (cellStrList.value(test_result_column) == "PASS")
            {
                if (failureList.size() != 0)
                {
                    int repass_index = -1;
                    for (int i = 0; i < failureList.size(); i++)
                    {
                        QStringList tempStrList = failureList.at(i);
                        if (tempStrList.value(0) == cellStrList.value(0))
                        {
                            repass_index = i;
                            break;
                        }
                    }

                    if (repass_index != -1)
                    {
                        failureList.removeAt(repass_index);
                    }

                }

                //  记录pass记录
                passStrList.append(cellStrList.value(0));   //  serial_id
                passStrList.append(cellStrList.value(6));   //  station
                QString test_result = cellStrList.value(test_result_column);
                QStringList test_result_str_list = test_result.split(":");

                passList.append(passStrList);

                continue;
            }

            failStrList.append(cellStrList.value(0));
            failStrList.append(cellStrList.value(6));
            QString test_result = cellStrList.value(test_result_column);



            QStringList test_result_str_list = test_result.split(":");
            if (test_result_str_list.size() > 2)
            {

                failStrList.append(((QString)test_result_str_list.at(1)).remove(0, 3));
                QString failResult =  ((QString)test_result_str_list.at(1)).remove(0, 3);

                for (QString failureCode : failureCodeList) {
                   if (failureCode.contains(failResult)) {
                       failStrList.append(failureCode);
                       QString failureDescription = failureCodeMap[failureCode].at(2);
                       failStrList.append(failureDescription);
                       QString failureMode = failureCodeMap[failureCode].at(3);
                       failStrList.append(failureMode);
                       QString repairGuidance = failureCodeMap[failureCode].at(4);
                       failStrList.append(repairGuidance);


                       break;
                   }
                }

                bool hasMoreErrorCode = test_result.contains(";");

                 //  有";".意味有多个错误(error code)
                int forTimes = 1;
                if (hasMoreErrorCode) {
                    forTimes = test_result.split(";").count();
                }

                test_result = test_result.remove(0, 5);
                QString errorCodeStr;
                for (int i = 0; i < forTimes; i++) {
                    QString tempErrorDesc = test_result.split(";").at(i);
                    QString errorCodeWithStationNum = QString(tempErrorDesc.split(":").at(0)).remove(0, 3);


                    QRegularExpression re("([A-Zsa-z]{2,3})-?([0-9]{2,3})");
                    QRegularExpressionMatch match = re.match(errorCodeWithStationNum);

                    if (match.hasMatch()) {
                            QString letters = match.captured(0);
//                            errorCodeStr = match.captured(2);
                            errorCodeStr.append(match.captured(2));
                            errorCodeStr.append("-");

                            qDebug() << "Letters:" << errorCodeStr;
                            qDebug() << "Numbers:" << "HI";
                        } else {
                            qDebug() << "The input string does not match the pattern.";
                        }
                }
                errorCodeStr.remove(errorCodeStr.size() - 1, 1);
                failStrList.append(errorCodeStr);

            }

            failureList.append(failStrList);


        }

        // Optional: Trim whitespace from each field
        //  标记内容是有效的,可以开始上方的统计
        for (QString& field : cellStrList) {
            field = field.trimmed();
            if (field.contains("serial_id"))
            {
                currentRow = true;
                continue;
            }
            if (field.contains("test_result"))
            {
                test_result_column = cellStrList.indexOf(field);
                break;
            }

        }

    }

    passNFailList["pass"] = passList;
    passNFailList["fail"] = failureList;
    qDebug() << failureList.size();
    file.close();

    return passNFailList;
}

int yantianamz::writeFile(QList<QStringList> allFailStrList,  QList<QStringList> allPassStrList)
{
    // 显示文件对话框，选择现有的Excel文件
//    QString filePath = QFileDialog::getOpenFileName(nullptr, "Open Excel File", "", "Excel Files (*.xls *.xlsx)");
//    if (filePath.isEmpty()) {
//        qDebug() << "No file selected.";
//        return -1;
//    }

    QString appInstalledPath = QCoreApplication::applicationDirPath();
    // 获取桌面文件夹的路径
    //    QString desktopPath = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation);
    appInstalledPath.append("/LT_offline_不良記錄_-_副本.xlsx");
    QFile outputFile (appInstalledPath);
    if (!outputFile.exists()) {
        return -1;
    }

    TWriteExcelThread* thread = new TWriteExcelThread(appInstalledPath, allFailStrList, allPassStrList,
                                                      importListMap, logDatetimeStr, projectName, failureCodeMap);
    QObject::connect(thread, &TWriteExcelThread::writedExcelFinished, this, [=](){
        this->tipsLabel->setText(tr("导出结束,请到程序目录查看"));
        this->selectBtn->setEnabled(true);
    });
    QObject::connect(thread, &TWriteExcelThread::writedExcelError, this, [=](const QString& errorMsg){
        QString errorFormatMsg = tr("导出异常:%1").arg(errorMsg);
        this->tipsLabel->setText(errorFormatMsg);
        this->selectBtn->setEnabled(true);
    });
    thread->start();
    return 1;
}

void yantianamz::writeExcel(const QList<QStringList> &allFailList)
{
//    // 创建一个 QXlsx::Document 对象
//    QXlsx::Document xlsx;

//    // 写入数据到单元格
//    xlsx.write("A1", "Hello Qt!");
//    xlsx.write("B1", 12345);
//    xlsx.write("C1", QDate(2023, 4, 1));

//    // 保存到文件
//    QString filePath = "example.xlsx";
//    if (xlsx.saveAs(outputFilePath)) {
//        qDebug() << "Excel file saved to" << filePath;
//    } else {
//        qDebug() << "Failed to save Excel file.";
//    }
       qDebug() << "hi";

}
yantianamz::~yantianamz()
{
    delete ui;
}

