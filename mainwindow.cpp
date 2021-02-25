#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QPushButton>
#include <iostream>

#include <QSqlDatabase>
#include <QSqlDriver>
#include <QSqlError>
#include <QSqlQuery>

#include <QDebug>

#include <QtWidgets/QApplication>
#include <QtWidgets/QMainWindow>
#include <QtCharts/QChartView>
#include <QtCharts/QBarSeries>
#include <QtCharts/QBarSet>
#include <QtCharts/QLegend>
#include <QtCharts/QBarCategoryAxis>
#include <QtCharts/QPieSlice>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

using namespace QXlsx;

using namespace std;
QT_CHARTS_USE_NAMESPACE

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    // database init
        DatabaseConnect();
        DatabaseInit();
        DatabasePopulate();



    // window title
    setWindowTitle("xlsx to SQL db demo");
    ui->setupUi(this);

    ChartInit();
}

MainWindow::~MainWindow()
{
    delete ui;
}

// db stuff
void MainWindow::DatabaseConnect()
{
    const QString DRIVER("QSQLITE");

    if(QSqlDatabase::isDriverAvailable(DRIVER))
    {
        QSqlDatabase db = QSqlDatabase::addDatabase(DRIVER);

        db.setDatabaseName(":memory:");

        if(!db.open())
            qWarning() << "MainWindow::DatabaseConnect - ERROR: " << db.lastError().text();

    }
    else
        qWarning() << "MainWindow::DatabaseConnect - ERROR: no driver " << DRIVER << " available";

}
void MainWindow::DatabaseInit()
{
    QSqlQuery query("CREATE TABLE product (id INTEGER PRIMARY KEY, name TEXT, value INTEGER)");

    if(!query.isActive())
        qWarning() << "MainWindow::DatabaseInit - ERROR: " << query.lastError().text();

}

//===== For demo purposes=====
void MainWindow::DatabasePopulate()
{
    QSqlQuery query;

    if(!query.exec("INSERT INTO product(name, value) VALUES('Bananas',54)"))
        qWarning() << "MainWindow::DatabasePopulate - ERROR: " << query.lastError().text();
    if(!query.exec("INSERT INTO product(name, value) VALUES('Apples',456)"))
        qWarning() << "MainWindow::DatabasePopulate - ERROR: " << query.lastError().text();
    if(!query.exec("INSERT INTO product(name, value) VALUES('Peaches',456)"))
        qWarning() << "MainWindow::DatabasePopulate - ERROR: " << query.lastError().text();
    if(!query.exec("INSERT INTO product(name, value) VALUES('Toothpaste',44)"))
        qWarning() << "MainWindow::DatabasePopulate - ERROR: " << query.lastError().text();
    if(!query.exec("INSERT INTO product(name, value) VALUES('Phones',87)"))
        qWarning() << "MainWindow::DatabasePopulate - ERROR: " << query.lastError().text();
}

//show db chart

void MainWindow::ChartInit()
{
    QSqlQuery query;
query.exec("SELECT name,value FROM product");
QList<QPieSlice *> ps;



    QPieSeries *series = new QPieSeries();

    while (query.next()) {

        QString name = query.value(0).toString();
        int value = query.value(1).toInt();

        QPieSlice *slice = new QPieSlice(name, value);
        ps.append(slice);
        series-> append(slice);

        qDebug() << name ;
    }

    for(auto x: ps){
        x ->setLabelVisible();
    }

    QChart *chart = new QChart();
       chart->addSeries(series);
       chart->setTitle("Pie chart from database");
       chart->setAnimationOptions(QChart::SeriesAnimations);
       chart->legend()->hide();

    ui->chartView->setChart(chart);
    ui -> chartView->setRenderHint(QPainter::Antialiasing);


}



// private slots
void MainWindow::on_button_open_file_clicked()
{
    cout << "Button clicked";
    QStringList fileNames = QFileDialog::getOpenFileNames(this, tr("Open File"),"/path/to/file/",tr("Excel files Files (*.xlsx)"));
    ui->listWidget->addItems(fileNames);

    readExcel(fileNames[0]);

}

void MainWindow::readExcel(QString path)
{
    // reading excel file
    Document xlsxR(path);
    if (xlsxR.load()) // load excel file
    {



        int sheetIndexNumber = 0;
        foreach( QString currentSheetName, xlsxR.sheetNames() )
        {
            // get current sheet
            AbstractSheet* currentSheet = xlsxR.sheet( currentSheetName );
            if ( NULL == currentSheet )
                continue;

            // get full cells of current sheet
            int maxRow = -1;
            int maxCol = -1;
            currentSheet->workbook()->setActiveSheet( sheetIndexNumber );
            Worksheet* wsheet = (Worksheet*) currentSheet->workbook()->activeSheet();
            if ( NULL == wsheet )
                continue;

            QString strSheetName = wsheet->sheetName(); // sheet name
            qDebug() << strSheetName;

            QVector<CellLocation> clList = wsheet->getFullCells( &maxRow, &maxCol );

            QVector< QVector<QString> > cellValues;
            for (int rc = 0; rc < maxRow; rc++)
            {
                QVector<QString> tempValue;
                for (int cc = 0; cc < maxCol; cc++)
                {
                    tempValue.push_back(QString(""));
                }
                cellValues.push_back(tempValue);
            }

            for ( int ic = 0; ic < clList.size(); ++ic )
            {
                CellLocation cl = clList.at(ic); // cell location

                int row = cl.row - 1;
                int col = cl.col - 1;

                QSharedPointer<Cell> ptrCell = cl.cell; // cell pointer

                // value of cell
                QVariant var = cl.cell.data()->value();
                QString str = var.toString();

                cellValues[row][col] = str;
            }



            for (int rc = 0; rc < maxRow; rc++)
            {


                for (int cc = 0; cc < maxCol; cc++)
                {
                    QString strCell = cellValues[rc][cc];
                    addToDB(cellValues[0][cc],cellValues[rc][cc].toInt());
                }
            }

            sheetIndexNumber++;
        }

    }
    ChartInit();
}

bool MainWindow::searchEntry(QString name)
{
QSqlQuery query;
query.prepare("select * from product where name = :name ");
query.bindValue(":name",name);
query.exec();

  while(query.next()){

          return true;
  }

  return false;
}

void MainWindow::addToDB(QString name, int valueX)
{
   QSqlQuery query;


if(searchEntry(name)){
  query.prepare("UPDATE product SET value = value + :value WHERE name =:name ");
  query.bindValue(":name",name);
  query.bindValue(":value",valueX);
  query.exec();

} else{

    query.prepare("INSERT INTO product (name, value) "
                    "VALUES (:name, :value)");
    query.bindValue(":name",name);
    query.bindValue(":value",valueX);
    query.exec();

    //if(!query.exec("INSERT INTO product(name, value) VALUES('"+name+"',"+valueX+"')"))
        //qWarning() << "MainWindow::addToDB- ERROR: " << query.lastError().text();
}

}
