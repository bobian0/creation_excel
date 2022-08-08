#ifndef MAINWINDOW_H
#define MAINWINDOW_H
#include <QMainWindow>
#include <QFileDialog>
#include <QVariant>
#include <QAxObject>
#include <QTime>
#include <QDebug>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    QTime time;

    int cellrow = 1;

private:
    Ui::MainWindow *ui;

private slots:
    void slot_pushbutton_clicked();
    void slot_exportToExcel();

};
#endif // MAINWINDOW_H
