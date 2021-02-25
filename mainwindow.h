#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QPushButton>


QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
private slots:

    void on_button_open_file_clicked();
    void readExcel(QString path);
    void addToDB(QString name, int value);
    bool searchEntry(QString name);


private:
    Ui::MainWindow *ui;
    QPushButton *button_open_file;


private:
    void DatabaseConnect();
    void DatabaseInit();
    void DatabasePopulate();
    void ChartInit();



};
#endif // MAINWINDOW_H
