#ifndef TWRITEEXCELTHREAD_H
#define TWRITEEXCELTHREAD_H

#include <QThread>
#include <QMap>


class TWriteExcelThread : public QThread
{
    Q_OBJECT
public:
    explicit TWriteExcelThread(const QString& outputPath
                               , const QList<QStringList>& allFailStrList
                               , const QList<QStringList>& allPassStrList
                               , const QMap<QString, QList<QStringList>>& importListMap
                               , const QString& dateStr
                               , const QString& projectName
                               , const QMap<QString, QStringList> failureCodeMap
                               , QObject *parent = nullptr);

protected:
    void run();
    void orderList();
    void writeSheetOne();
    void writeSheetTwo();

private:
    QList<QStringList> m_allPassStrList;
    QList<QStringList> m_allFailStrList;
    QMap<QString, QList<QStringList>> m_importListMap;
    QString m_outputPath;
    QString m_dateStr;
    QString m_projectName;

    QList<QStringList> allStationFailList;
    QList<QStringList> allStationPassList;
    QMap<QString, QStringList> m_failureCodeMap;
    QStringList failureCodeKeyStrList;

signals:
    void writedExcelFinished();
    void writedExcelError(const QString& errorMsg);
};

#endif // TWRITEEXCELTHREAD_H
