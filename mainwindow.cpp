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

////    QAxObject *word = connectWord();
////    QAxObject *documents = createDocuments(word);
////    QAxObject *document = createDocument(documents,"D:\\testdot.dot"); //добавляем свой документ в коллекцию
////     wordOptions(word,false,true);

//QAxObject *word = new QAxObject("Word.Application", this);
//word->setProperty("DisplayAlerts", false);
//word->setProperty("Visible", true);
//QAxObject *documents = word->querySubObject("Documents"); //получаем коллекцию документов
//QAxObject *document = documents->querySubObject("Add(D:\\testdot.dot)"); //
//QAxObject *document1 = documents->querySubObject("Add()");

//activeDocument(document);

//////поиск по индексу
////QString c = documents->dynamicCall("count").toString();
////QAxObject *item = documents->querySubObject("Item(const QVariant &)", 2);
////QString name = item->dynamicCall("FullName").toString();
//////конец поиска по индексу
//////закрываю документ по индексу
////item->dynamicCall("Close(wdDoNotSaveChanges)");
//bool c ;
//c = checkAndCloseDocument(documents, "testdot", false);
//c = checkAndCloseDocument(documents, "Документ2", false);
//   // word->querySubObject("ActiveDocument")->dynamicCall("Close()");
// // document1->querySubObject("SaveAs()", "D:\\test.docx");
//   // disconnectWord(word);

  ActiveWord word;

 QAxObject* document1 = word.documentOpen(true);
 QAxObject* document2 = word.documentOpen(false);

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
