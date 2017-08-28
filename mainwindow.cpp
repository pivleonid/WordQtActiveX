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


  ActiveExcel excel;
  excel.documentOpen("D:\\testil.xlsx");
  QAxObject* sheet = excel.documentSheetActive("Лист1");
  excel.sheetCellPaste(sheet, "hi",1,1); //запись в ячейку A1

  excel.documentAddSheet();
  sheet = excel.documentSheetActive("Лист2");
  excel.sheetCellPaste(sheet, "hi1",1,1); //запись в ячейку A1

  QVariant a = excel.sheetCellInsert(sheet, 1 ,1); //"a" хранит значение ячейки листа 2
  sheet = excel.documentSheetActive("Лист1"); //переключаемся на лист1
  excel.sheetCopyToBuf(sheet, "B2:C16"); // копирование в буфер

  ActiveWord word;
  word.documentOpen(true, "D:\\template1.docx"); //метка label1
  word.selectionPasteTextFromBuffer("[label1]");




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
