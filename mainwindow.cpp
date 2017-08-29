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


//  ActiveExcel excel;
//  QAxObject* workbook = excel.documentOpen("D:\\testil.xlsx");
//  QAxObject* sheet = excel.documentSheetActive("Лист1");
//  excel.sheetCellPaste(sheet, "hi",1,1); //запись в ячейку A1
//  excel.documentAddSheet("Лист-хуист");
//  excel.documentAddSheet();
//  sheet = excel.documentSheetActive("Лист-хуист");
//  excel.sheetCellPaste(sheet, "hi111",1,1); //запись в ячейку A1
//  //excel.sheetProperty("name13", workbook);

//  excel.documentAddSheet();
//  sheet = excel.documentSheetActive("Лист2");
//  excel.sheetCellPaste(sheet, "hi1",1,1); //запись в ячейку A1

//  QVariant a = excel.sheetCellInsert(sheet, 1 ,1); //"a" хранит значение ячейки листа 2
//  sheet = excel.documentSheetActive("Лист1"); //переключаемся на лист1
//  excel.sheetCopyToBuf(sheet, "B2:C16"); // копирование в буфер

//  ActiveWord word;
//  word.documentOpen(true, "D:\\template1.docx"); //метка label1
//  word.selectionPasteTextFromBuffer("[label1]");

//  ActiveWord word;
//   QAxObject* doc1 = word.documentOpen(true, "D:/Otpuck3.docx");
//   word.selectionFindReplaseAll("[code1]", "1286", false);
//   word.selectionFindReplaseAll("[code2]", "720-816", false);
//   word.selectionFindReplaseAll("[nameOrganization]", "НПП ГАММА", false);
//   word.selectionFindReplaseAll("[post]", "Руководитель отдела разработки ПО", false);
//   word.selectionFindReplaseAll("[numberDoc]", "1", false);
//   word.selectionFindReplaseAll("[date]", "12.07.1993", false);
//   word.selectionFindReplaseAll("[year]", "17", false);


//   QAxObject* doc2 = word.documentOpen(true, "D:/Otpuck4.docx");

//   QStringList list = word.tableGetLabels(1, 2);

//   QStringList label;
//   label << "[1]"<<"[2]"<<"[3]"<<"[4]"<<"[5]"<<"[6]"<<"[7]"<<"[8]"<<"[9]";

//   QList<QStringList> table;
//   QStringList table1;
//   table1.append("ПКД");
//   table1.append("Инженер");
//   table1.append("Кирьянов И.О.");
//   table1.append( "1");
//   table1.append("12");
//   table1.append("12.07.1999");
//   table1.append("12.07.1999");
//   table1.append( "документ №3");
//   table1.append( "12.08.1999");
//   table.append(table1);
// //  //
//    QStringList table2;
//   table2.append("ПК1Д");
//   table2.append("Инженер1");
//   table2.append("Кирьяноasdв И.О.");
//   table2.append( "112");
//   table2.append("132");
//   table2.append("12.07.1999");
//   table2.append("12.07.1999");
//   table2.append( "документ №3");
//   table2.append( "12.08.1999");
//   table.append(table2);


//   word.tableFill(table,label,1, 3); //1 таблица 3 строка

//   word.selectionFindAndPasteBuffer(doc2,doc1, "[lable_tab1]" );

//  word.documentSave(doc1,"D:/", "otpuskFull", "docx");

  ActiveWord word;
   word.documentOpen(true, "D:\\Freq.docx"); //метка label1
  word.tableMergeCell(1, "[up_mkv]","Яблочки", 1, 1);

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
