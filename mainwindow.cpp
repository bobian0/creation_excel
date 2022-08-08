#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    time = QTime::currentTime();

    connect(ui->pushbutton_2,SIGNAL(clicked()),this,SLOT(slot_pushbutton_clicked()));
    connect(ui->pushbutton,SIGNAL(clicked()),this,SLOT(slot_exportToExcel()));
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::slot_exportToExcel()
{
    //保存文件的路径
       QString filepath = QFileDialog::getSaveFileName(this, tr("创建文件"), time.toString("hh_mm_ss"));
       if(!filepath.isEmpty())
       {
           QAxObject *excel = new QAxObject(this);
           excel->setControl("Excel.Application");
           excel->dynamicCall("SetVisible (bool Visible)","false");
           excel->setProperty("DisplayAlerts", false);
           QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
           workbooks->dynamicCall("Add");//新建一个工作簿
           QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
           QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
           QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);
           QAxObject *cellA, *cellB, *cellC;
           //设置标题
           int cellrow = 1;
           QString A = "A" + QString::number(cellrow);
           QString B = "B" + QString::number(cellrow);
           QString C = "C" + QString::number(cellrow);
           cellA = worksheet->querySubObject("Range(QVariant, QVariant)", A);
           cellB = worksheet->querySubObject("Range(QVariant, QVariant)", B);
           cellC = worksheet->querySubObject("Range(QVarinat, QVariant)", C);
           //设置单位的值
           cellA->dynamicCall("SetValue(const QVariant&)", QVariant("姓名"));
           cellB->dynamicCall("SetValue(const QVariant&)", QVariant("学号"));
           cellC->dynamicCall("SetValue(const QVariant&)", QVariant("角度值"));
           cellrow++;

           worksheet->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
           workbook->dynamicCall("Close()");
           excel->dynamicCall("Quit()");
           delete excel;
           excel = nullptr;
       }
}

//添加数据
void MainWindow::slot_pushbutton_clicked()
{
    qDebug() << "start";
    cellrow++;

    QString filepath = QFileDialog::getSaveFileName(this, tr("Save orbit"), ".", tr("Microsoft Office 2019 (*xls, *.xlsx)"));
    if(!filepath.isEmpty())
    {

        QAxObject *excel = new QAxObject(this);
        excel->setControl("Excel.Application");
        excel->dynamicCall("SetVisible (bool Visible)","false");
        excel->setProperty("DisplayAlerts", false);
        QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
        workbooks->dynamicCall("Add");//新建一个工作簿
        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
        QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
        QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);

        //设置标题
        QAxObject * Range1 =worksheet->querySubObject("Range(QString)","A1");
        Range1->setProperty("Value","姓名");
        QAxObject * Range2 =worksheet->querySubObject("Range(QString)","B1");
        Range2->setProperty("Value","学号");
        QAxObject * Range3 =worksheet->querySubObject("Range(QString)","C1");
        Range3->setProperty("Value","角度值");

        //插入信息
        QAxObject *cellA, *cellB, *cellC;
        QString A = "A" + QString::number(cellrow);
        QString B = "B" + QString::number(cellrow);
        QString C = "C" + QString::number(cellrow);
        cellA = worksheet->querySubObject("Range(QVariant, QVariant)", A);
        cellB = worksheet->querySubObject("Range(QVariant, QVariant)", B);
        cellC = worksheet->querySubObject("Range(QVarinat, QVariant)", C);
        //插入数据
//        cellA->dynamicCall("SetValue(const QVariant&)", QVariant(ui->lineedit_name->text()));
//        cellB->dynamicCall("SetValue(const QVariant&)", QVariant(ui->lineedit_id->text()));
//        cellC->dynamicCall("SetValue(const QVariant&)", QVariant(QString::number(-30 + rand()%60)));

        for(int i = 0; i < 50; i++)
        {
            //创建三块空间（excel的列数）
//            QString A = "A" + QString::number(cellrow);
//            QString B = "B" + QString::number(cellrow);
//            QString C = "C" + QString::number(cellrow);
//            cellA = worksheet->querySubObject("Range(QVariant, QVariant)", A);
//            cellB = worksheet->querySubObject("Range(QVariant, QVariant)", B);
//            cellC = worksheet->querySubObject("Range(QVarinat, QVariant)", C);
            ++cellrow;
            //向空间内插入内容
            cellA->dynamicCall("SetValue(const QVariant&)", QVariant(ui->lineedit_name->text()));
            cellB->dynamicCall("SetValue(const QVariant&)", QVariant(ui->lineedit_id->text()));
            cellC->dynamicCall("SetValue(const QVariant&)", QVariant(QString::number(-30 + rand()%60)));


        }

        qDebug() << cellrow;

        worksheet->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
        workbook->dynamicCall("Close()");
        excel->dynamicCall("Quit()");
        delete excel;
        excel = nullptr;

    }
}

