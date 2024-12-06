#include "yantianamz.h"

#include <QApplication>

#include <QLocale>
#include <QTranslator>
#include <QDebug>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    // 检测系统语言
    QLocale systemLocale = QLocale::system();
    QString languageCode = systemLocale.name().left(2); // 获取语言代码，如"en"、"zh"等
    if (languageCode == "en") {
        languageCode = "English";
    }
    else {
        languageCode = "Chinese";
    }

    // 创建翻译器对象
    QTranslator translator;

    // 根据语言代码加载相应的翻译文件
    // 假设你的翻译文件命名为"YourApp_xx.qm"，其中"xx"是语言代码
    QString translationFile = QString("lang_%1.qm").arg(languageCode);
    if (!translator.load(":/i18n/lang_Chinese.qm")) {
        qWarning() << "Failed to load translation file:" << translationFile;
    } else {
        // 应用翻译
        a.installTranslator(&translator);
    }

    yantianamz w;
    w.show();
    return a.exec();
}
