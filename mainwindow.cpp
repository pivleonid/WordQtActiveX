#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "qaxobject.h"
  #include "activeword.h"
#include "qdebug.h"

#include <windows.h>



MainWindow::MainWindow(QWidget *parent) :
  QMainWindow(parent),
  ui(new Ui::MainWindow)
{
  ui->setupUi(this);



//  QAxObject *word = new QAxObject("Word.Application", this);
//  Sleep(1000);
//  word->setProperty("DisplayAlerts", false);
//  Sleep(1000);
//  word->setProperty("Visible", true);
//  Sleep(1000);
//  QAxObject *documents = word->querySubObject("Documents"); //получаем коллекцию документов
//  QAxObject *document = documents->querySubObject("Add(D:\\tabl.docx)");
////-------------
//  QAxObject* wordSelection = word->querySubObject("Selection");
//  wordSelection->dynamicCall("WholeStory()");
//   QList<QVariant> params;//Все параметры не обязательные!
//   params.operator << (QVariant(1));//[Separator]
//   params.operator << (QVariant(2));//[NumRows]
//   params.operator << (QVariant(3));//[NumColumns]
//   params.operator << (QVariant(false));// [InitialColumnWidth]
//   //
//   params.operator << (QVariant(0));      //[Format]
//   params.operator << (QVariant(true));   //  [ApplyBorders]
//   params.operator << (QVariant(false));  //[ApplyShading]
//   params.operator << (QVariant(true));   //[ApplyFont]
//   params.operator << (QVariant(true));   //[ApplyColor]
//   //
//   params.operator << (QVariant(true));   //[ApplyHeadingRows]
//   params.operator << (QVariant(false));  //[ApplyLastRow]
//   params.operator << (QVariant(true));   // [ApplyFirstColumn]
//   //
//   params.operator << (QVariant(false));  //[ApplyLastColumn]
//   params.operator << (QVariant(true));   //[AutoFit]
//   params.operator << (QVariant(1));      //[AutoFitBehavior]
//   params.operator << (QVariant(1));      //[DefaultTableBehavior]
//   QVariant param;

//   param =    wordSelection->dynamicCall("ConvertToTable(const QVariant&,const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&, const QVariant&)", params);


  ActiveWord word;
  word.documentOpen(true, "D:\\Freq.docx");
  //
  QStringList label;
  label  <<"[freq_mhz]"<< "[usp_db]"<< "[up_mkv]"<< "[test]"<<"[up_db]"<<"[usp_mkv]";
  //
  QList<QStringList> table;
  for( uint i =0 ; i < 6; i++){
      QStringList temp;
      for(uint j = 0; j < 6; j++)
        temp.append( QString::number(j) );
      table.append(temp);
  }
  word.tableFill(table,label,1);



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
