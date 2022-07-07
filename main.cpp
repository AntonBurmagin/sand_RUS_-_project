#include <QApplication>
#include <QPushButton>
#include <QAxObject>
#include <QAxWidget>
#include <QFileDialog>
#include <iostream>
#include <QDebug>
#include <QLabel>
#include <string>
#include <vector>
#include <QPainter>
#include <QPaintEvent>
#include <QPicture>
#include <QComboBox>
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QFont>
#include <QSpinBox>
#include <QCheckBox>
#include <QLineEdit>


struct sandMap {
    QString name;
    std::vector<double> vec;
};


int main(int argc, char *argv[]) {
    QApplication a(argc, argv);

    std::vector<double> xs = {0, 0.05, 0.16, 0.315, 0.63, 1.25, 2.5, 5, 10, 20};    // значения x для графика по гост 8736
    std::vector<sandMap> vector;        // карты намыва
    QVector<QWidget *> maps;

    auto window = new QWidget(nullptr);
    QVBoxLayout* vbox = new QVBoxLayout(window);
    QHBoxLayout hbox(nullptr);
    QPushButton button("file");
    hbox.addWidget(&button);
    QLabel label(nullptr);
    QComboBox sheet(nullptr);
    hbox.addWidget(&sheet);
    QCheckBox standard_8736(nullptr);
    standard_8736.setText("ГОСТ 8736-2014");
    standard_8736.setCheckState(Qt::Unchecked);
    QCheckBox standard_25100(nullptr);
    standard_25100.setText("ГОСТ 25100-2011");
    standard_25100.setCheckState(Qt::Unchecked);
    hbox.addWidget(&standard_8736);
    hbox.addWidget(&standard_25100);
    QAxWidget excel("Excel.Application");
    sheet.setDisabled(true);
    sheet.setPlaceholderText("№");
    label.setSizePolicy(QSizePolicy::Expanding,QSizePolicy::Expanding);
    label.setWordWrap(true);

    label.setText("Choose file");
    label.setFont(QFont("Arial", 40, 3));
    label.setStyleSheet("QLabel {background-color: lightblue; color: red;}");
    label.setAlignment(Qt::AlignCenter);
    vbox->addLayout(&hbox);
    vbox->addWidget(&label);
    QPushButton start("start");
    start.setDisabled(true);
    vbox->addWidget(&start);

    QString path = a.applicationFilePath();
    auto file = new QFileDialog();
    file->setDirectory(path);
    QWidget::connect(&button, &QPushButton::clicked, [file](){
        file->open();
    });
    QWidget::connect(file, &QFileDialog::fileSelected,[&label, &sheet, &excel, &start](auto filePath) {
        excel.setProperty("Visible", false);
        QAxObject *workbooks = excel.querySubObject("WorkBooks");
        workbooks->dynamicCall("Open (const QString&)", filePath);
        QAxObject *workbook = excel.querySubObject("ActiveWorkBook");
        int countSheets = workbook->querySubObject("WorkSheets")->property("Count").toInt();
        if(sheet.count()>0) {
            sheet.clear();
            label.setStyleSheet("QLabel {background-color: lightblue; color: red;}");
            //start.setDisabled(true);
        }
        for (int n = 1; n <= countSheets; n++) {
            sheet.addItem(QString::number(n));
        }
        label.setText("Choose number of sheet with initial data and standard for result");
        sheet.setEnabled(true);
    });
    QObject::connect(&standard_8736, &QCheckBox::clicked, [&standard_25100, &sheet, &start, &label](auto state){
        if(state==true) {
            standard_25100.setCheckState(Qt::Unchecked);
            if (!sheet.currentText().isNull()) {
                start.setEnabled(true);
                label.setText("Press start");
                label.setStyleSheet("QLabel {background-color: lightblue; color: green;}");
            }
        } else {
            start.setEnabled(false);
            label.setText("Choose number of sheet with initial data and standard for result");
            label.setStyleSheet("QLabel {background-color: lightblue; color: red;}");
        }
    });
    QObject::connect(&standard_25100, &QCheckBox::clicked, [&standard_8736, &sheet, &start, &label](auto state){
        if(state==true) {
            standard_8736.setCheckState(Qt::Unchecked);
            if (!sheet.currentText().isNull()) {
                start.setEnabled(true);
                label.setText("Press start");
                label.setStyleSheet("QLabel {background-color: lightblue; color: green;}");
            }
        } else {
            start.setEnabled(false);
            label.setText("Choose number of sheet with initial data and standard for result");
            label.setStyleSheet("QLabel {background-color: lightblue; color: red;}");
        }
    });
    QObject::connect(&sheet, &QComboBox::currentTextChanged, [&start, &standard_8736, &standard_25100, &label]() {
        if(standard_25100.isChecked() || standard_8736.isChecked()) {
            start.setEnabled(true);
            label.setText("Press start");
            label.setStyleSheet("QLabel {background-color: lightblue; color: green;}");
        }
    });

    QWidget::connect(&start, &QPushButton::clicked, [&maps](){
        for (auto map : maps) {
            map->close();
        }
        maps.clear();
        std::cout << "Maps size: " << maps.size() <<std::endl;
    });


    QObject::connect(&start, &QPushButton::clicked, [&excel, &file, &sheet, &label, &xs, &vector, &maps](){
        if (!vector.empty())
            vector.clear();

        int rows;
        int columns;
        auto filesHistory = file->selectedFiles();
        QString filePath = filesHistory[filesHistory.size()-1];
        QAxObject *workbooks = excel.querySubObject("WorkBooks");
        workbooks->dynamicCall("Open (const QString&)", filePath);
        QAxObject *workbook = excel.querySubObject("ActiveWorkBook");
        QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", sheet.currentText().toInt());
        QAxObject *range = worksheet->querySubObject("UsedRange");
        auto startRow = range->property("Row").toInt();
        auto startCol = range->property("Column").toInt();
        QAxObject* rowsQAx = range->querySubObject("Rows");
        QAxObject* columnsQAx = range->querySubObject("Columns");
        rows = rowsQAx->dynamicCall("Count").toInt();
        columns = columnsQAx->dynamicCall("Count").toInt();



        for(int i=startRow; i<30; i++){
            for (int j=startCol; j<2; j++) {
                QAxObject *cell = worksheet->querySubObject("Cells( , )", i, j);
                auto str = cell->property("Value");
                if (str.toString().toLower().contains("карта")) {
                    sandMap add;
                    add.name = str.toString();
                    double multiplier = (2000 -
                            worksheet->querySubObject("Cells( , )", i-1, 2)->property("Value").toDouble() -
                            worksheet->querySubObject("Cells( , )", i-1, 3)->property("Value").toDouble())/1000;
                    add.vec.push_back(0);
                    for (int k=11; k>1; k--) {
                        if (k==11) {
                            add.vec.push_back((worksheet->querySubObject("Cells( , )", i - 1, k-2)              // вставка значения в граммах под ситом 0.05
                                    ->property("Value").toDouble()
                                * multiplier)
                                * (worksheet->querySubObject("Cells( , )", i-1, k)
                                    ->property("Value").toDouble()
                                / worksheet->querySubObject("Cells( , )", i, k-2)
                                    ->property("Value").toDouble()));
                            add.vec.push_back((worksheet->querySubObject("Cells( , )", i - 1, k-2)              // вставка значения в граммах под ситом 0.16
                                ->property("Value").toDouble()
                                * multiplier) - add.vec[add.vec.size()-1]);
                            k -= 2;
                        } else if (k>3){
                            add.vec.push_back(worksheet->querySubObject("Cells( , )", i - 1, k)               // вставка значения в граммах под ситами до 2.5
                                ->property("Value").toDouble()
                            * multiplier);
                        } else {
                            add.vec.push_back(worksheet->querySubObject("Cells( , )", i - 1, k)                // вставка значения в граммах под ситами 5 и 10
                                                      ->property("Value").toDouble());
                        }
                    }
                    double weightSum = 0;
                    for (auto weight : add.vec)
                        weightSum += weight;
                    for (int i=0; i<add.vec.size(); i++) {                                  // перевод в % накопительным итогом
                        if (i!=0)
                            add.vec[i] = (add.vec[i] / weightSum) * 100 + add.vec[i-1];
                        else
                            add.vec[i] = (add.vec[i] / weightSum) * 100;
                    }
                    vector.push_back(add);
                    for (auto percent : add.vec)
                        std::cout << percent << " ";
                    std::cout << std::endl;
                }
            }
        }
        if (vector.size()>0) {
            QPen penForMap(Qt::darkGreen, 5);


            int loop = 0;
            for (auto i : vector) {
                auto wid = new QWidget(nullptr);
                maps.push_back(wid);
                QVBoxLayout* verticalBox = new QVBoxLayout(wid);
                QHBoxLayout *horizontalBox = new QHBoxLayout(nullptr);
                QLineEdit* x = new QLineEdit(nullptr);
                x->setPlaceholderText("input x");
                QLineEdit* y = new QLineEdit(nullptr);
                y->setPlaceholderText("result y");
                y->setDisabled(true);
                horizontalBox->addWidget(x);
                horizontalBox->addWidget(y);

                QWidget::connect(x, &QLineEdit::textChanged, [y, &xs, &vector, &loop](auto text) {  // нахождение значения У от Х
                    double valueX = text.toDouble();
                    if (0<=valueX && valueX<=20) {
                        for (int t=0; t<xs.size(); t++){
                            if (xs[t]>valueX) {
                                double valueY = vector[0].vec[t - 1] +
                                        (vector[0].vec[t] - vector[0].vec[t - 1]) *
                                            ((valueX - xs[t - 1]) / (xs[t] - xs[t - 1]));
                                y->setText(QString::number(valueY));
                                break;
                            }
                        }
                    } else if(valueX>20) {
                        y->setText(QString::number(100));
                    } else {
                        y->setText("Incorrect X");
                    }
                });

                QLabel* mapGraphic= new QLabel(nullptr);
                verticalBox->addLayout(horizontalBox);
                verticalBox->addWidget(mapGraphic);
                maps[loop]->resize(600, 300);
                mapGraphic->setScaledContents(true);

                QPicture graphic;
                QPainter painter(&graphic);
                painter.setPen(penForMap);
                maps[loop]->setWindowTitle(i.name);
                float factorX = maps[loop]->width() / 10;
                float factorY = maps[loop]->height() / 100;
                for (int k = 0; k < (i.vec.size()-1); k++)
                    painter.drawLine(k * factorX, 600 - (i.vec[k] * factorY),
                                     (k + 1) * factorX, 600 - (i.vec[k + 1]) * factorY);
                painter.end();
                mapGraphic->setPicture(graphic);
                maps[loop]->move(loop * 100, 100 * (loop +1));
                loop++;
            }
            for (auto i : maps)
                i->show();
        } else {
            label.setText("No maps in file\nTry other file or sheet");
            label.setStyleSheet("QLabel {background-color: lightblue; color: red;}");
            sheet.setCurrentText("№");
        }
        excel.dynamicCall("Quit (void)");
    });


    window->setWindowTitle("Main");
    window->resize(800, 500);
    window->move(0,0);
    window->show();

    /* Сохранение в ЭКСЕЛЬ!!!
    worksheet->querySubObject("Cells( , )",1 ,1)->setProperty("Value", "Привет");
    excel.setProperty("Visible", true);
    workbook->dynamicCall("SaveAs (const QString&)",
                           "C:\\Users\\burma\\Desktop\\C++\\HW\\HW-40 unit tests\\untitled\\1ex.xlsx");
    */






    return QApplication::exec();
}
