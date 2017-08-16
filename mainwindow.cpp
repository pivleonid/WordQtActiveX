#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "qaxobject.h"
  #include "activeword.h"
#include "activeexcel.h"
#include "qdebug.h"

#include <windows.h>



MainWindow::MainWindow(QWidget *parent) :
  QMainWindow(parent),
  ui(new Ui::MainWindow)
{
  ui->setupUi(this);


//  ActiveExcel Excel;
//  QAxObject* doc1 = Excel.documentOpen();
//  QAxObject* sheet1 = Excel.documentAddSheet(doc1);

//   QAxObject *statSheet  =  Excel.documentSheetActive( "Лист1");
//  Excel.sheetCellPaste(statSheet, "hi",1,1);
// statSheet  =  Excel.documentSheetActive( "Лист2");
// Excel.sheetCellPaste(statSheet, "hi",2,1);

// QVariant a = Excel.sheetCellInsert(statSheet, 2 ,1);






  QAxObject *mExcel = new QAxObject( "Excel.Application",this);
  mExcel->setProperty("DisplayAlerts", false);
  mExcel->setProperty("Visible", true);
  // на книги
  QAxObject *workbooks = mExcel->querySubObject( "Workbooks" );
  // на директорию, откуда грузить книгу
  //QAxObject *workbook = workbooks->querySubObject( "Add" );
  QAxObject *workbook = workbooks->querySubObject( "Add(const QVariant&)" , "D:\\testil.xlsx" );
  // на листы (снизу вкладки)
  QAxObject *mSheets = workbook->querySubObject( "Sheets" );
  // указываем, какой лист выбрать. У меня он называется topic.
  QAxObject* mSheets1 = mSheets->querySubObject("Add");
  //указатель на нужный лист
  QAxObject *StatSheet = mSheets->querySubObject( "Item(const QVariant&)", QVariant("Лист1") );
  // получение указателя на ячейку [row][col] ((!)нумерация с единицы)
  QAxObject* cell = StatSheet->querySubObject("Cells(QVariant,QVariant)", 1 , 1);
  // вставка значения переменной data (любой тип, приводимый к QVariant) в полученную ячейку
  cell->setProperty("Value", "Hello");
  //
  StatSheet = mSheets->querySubObject( "Item(const QVariant&)", QVariant("Лист1") );
  // получение указателя на ячейку [row][col] ((!)нумерация с единицы)
  cell = StatSheet->querySubObject("Cells(QVariant,QVariant)", 1 , 2);
  // вставка значения переменной data (любой тип, приводимый к QVariant) в полученную ячейку
  cell->setProperty("Value", "Hello1");

  //вытаскиваю значения из ячеек


///ширина столбцов
  QAxObject *rangec = StatSheet->querySubObject( "Range(const QVariant&)",QVariant("D2:E6"));
// получаю указатель на строку
QAxObject *razmer = rangec->querySubObject("Rows");
// устанавливаю её размер.
razmer->setProperty("RowHeight",68);
razmer = rangec->querySubObject("Columns");
// устанавливаю её размер.
razmer->setProperty("ColumnWidth",34);


  delete mExcel;

  int i;
  i++;



}
MainWindow::~MainWindow()
{
  delete ui;
}



/*

 Selection.TypeText Text:="ваыаываываываываыв"
    Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
    Selection.TypeText Text:="ываываываываываываыв"
    Windows("Документ1").Activate
    Selection.TypeParagraph
    Selection.TypeText Text:="ываываываываываываываы"
    Windows("Документ2").Activate
    Windows("Документ1").Activate
    Windows("Документ2").Activate
    Selection.TypeText Text:="    фывыфвыфвфывфывфывф"
 */
/*Набор юного тестировщика
  QAxObject *word = new QAxObject("Word.Application", this);
  Sleep(1000);
  word->setProperty("DisplayAlerts", false);
  Sleep(1000);
  word->setProperty("Visible", true);
  Sleep(1000);
  QAxObject *documents = word->querySubObject("Documents"); //получаем коллекцию документов
  QAxObject *document = documents->querySubObject("Add()");*/
