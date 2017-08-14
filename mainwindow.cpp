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





  ActiveWord word;
  word.documentOpen(true, "D:\\Otpuck3.docx");
  word.selectionFindReplaseAll("[code1]", "1286", false);
  word.selectionFindReplaseAll("[code2]", "720-816", false);
  word.selectionFindReplaseAll("[nameOrganization]", "НПП ГАММА", false);
  word.selectionFindReplaseAll("[post]", "Руководитель отдела разработки ПО", false);
  word.selectionFindReplaseAll("[numberDoc]", "1", false);
  word.selectionFindReplaseAll("[date]", "12.07.1993", false);
  word.selectionFindReplaseAll("[year]", "17", false);

  QStringList list = word.tableGetLabels(5, 2);

  QStringList label;
  label << "[1]"<<"[2]"<<"[3]"<<"[4]"<<"[5]"<<"[6]"<<"[7]"<<"[8]"<<"[9]";

  QList<QStringList> table;
  QStringList table1;
  table1.append("ПКД");
  table1.append("Инженер");
  table1.append("Кирьянов И.О.");
  table1.append( "1");
  table1.append("12");
  table1.append("12.07.1999");
  table1.append("12.07.1999");
  table1.append( "документ №3");
  table1.append( "12.08.1999");
  table.append(table1);
//  //
   QStringList table2;
  table2.append("ПК1Д");
  table2.append("Инженер1");
  table2.append("Кирьяноasdв И.О.");
  table2.append( "112");
  table2.append("132");
  table2.append("12.07.1999");
  table2.append("12.07.1999");
  table2.append( "документ №3");
  table2.append( "12.08.1999");
  table.append(table2);


  word.tableFill(table,label,5, 2);
  int j = 1;
  j++;

  //  ActiveWord word;
  //  word.documentOpen(true, "D:\\Freq.docx");
  //  //
  //  QStringList label;
  //  label  <<"[freq_mhz]"<< "[usp_db]"<< "[up_mkv]"<< "[test]"<<"[up_db]"<<"[usp_mkv]" << "[Hello1]"
  //           << "[Hello2]"<< "[Hello3]"<< "[Hello4]"<< "[Hello5]";//<<"[usp_mkv]";
  //  //
  //  QList<QStringList> table;
  //  for( uint i =0 ; i < 6; i++){
  //      QStringList temp;
  //      for(uint j = 0; j < 6; j++)
  //        temp.append( QString::number(j) );
  //      table.append(temp);
  //  }
  //  word.tableFill(table,label,1);


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
