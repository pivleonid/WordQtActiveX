#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "qaxobject.h"
  #include "activeword.h"
#include "activeexcel.h"
#include "qdebug.h"

#include <windows.h>
#include <qmessagebox.h>


MainWindow::MainWindow(QWidget *parent) :
  QMainWindow(parent),
  ui(new Ui::MainWindow)
{
  ui->setupUi(this);


//  QStringList strListNamelabel;
//  strListNamelabel << "[Устройства]" << "[Конденсаторы]"<<"[Микросхемы]"<<"[Светодиоды]"<<"[Дроссели]"<<"[Резисторы]"<<"[Коммутация]"<<"[Диоды]"<<"[Транзисторы]"<<"[Контактные соединения]"<<"[Фильтры]"<<"[Кварцевый резонатор]"<<"[Предохранители]";

//  ActiveWord word;
//  if(!word.wordConnect()){
//      QMessageBox msgBox;
//      msgBox.setText("Word не установлен");
//      msgBox.exec();
//      return;
//    }
// // QString path = QApplication::applicationDirPath() + "/center.docx";

//  QString path = "D:/projects/WordQtActiveX-master/center.docx";

//  QAxObject* doc1 = word.documentOpen(path);
//  if(doc1 == NULL){
//      QMessageBox msgBox;
//        msgBox.setText("Не найден шаблон");
//        msgBox.exec();
//    return;
//    }
//   word.setVisible();
//   word.tableSizeRows(1,1);
//  foreach (QString var, strListNamelabel) {
//      //подчеркивание
//      QVariant a = word.selectionFindFontname(var, true, false, true, true, "GOST type B");
//      //центрирование
//      word.selectionAlign(var , false, false, true);
//      QString s = var;
//      s.remove(0,1);
//      s.remove(s.count()-1,1);
//      QString s1 = var;
//      // замена меток/
//      word.findReplaseLabel(s1, s, true);
//      s.clear();
//      s1.clear();
//    }

  ActiveExcel excel;
  bool tr = excel.excelConnect();
  QAxObject* workbook = excel.workbookOpen("D:\\testil.xlsx");
  QAxObject* sheet = excel.workbookSheetActive("Лист1");
  QVariant dataG, dataY, dataR, data_, dataB;
   excel.setVisible(true);
  excel.sheetCellColorInsert(sheet, dataG, 2, 2);
  excel.sheetCellColorInsert(sheet, dataY, 3, 2);
  excel.sheetCellColorInsert(sheet, dataR, 4, 2);
  excel.sheetCellColorInsert(sheet, data_, 5, 2);
  excel.sheetCellColorInsert(sheet, dataB, 6, 2);

  excel.workBookClose(workbook);

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
