#include "activeexcel.h"

ActiveExcel::ActiveExcel()
{
  excelApplication_ = new QAxObject( "Excel.Application");
  excelApplication_->setProperty("DisplayAlerts", false);
  excelApplication_->setProperty("Visible", true);
  documents_ = excelApplication_->querySubObject( "Workbooks" );



}
ActiveExcel::~ActiveExcel(){
  excelApplication_->dynamicCall("Quit()");
  delete sheets_;
  delete documents_;
  delete excelApplication_;
}


QAxObject* ActiveExcel::documentOpen(QVariant path){
  if (path == "")
    return documents_->querySubObject("Add");
  return documents_->querySubObject("Add(const QVariant &)", path);

}

QAxObject* ActiveExcel::documentAddSheet(QAxObject* document, QVariant sheet ){

  sheets_ = document->querySubObject("Sheets");
  if (sheet == "")
   return   sheets_->querySubObject("Add");
}

QAxObject* ActiveExcel::documentSheetActive(QAxObject* sheet1, QVariant sheet){
    return sheet1->querySubObject( "Item(const QVariant&)", sheet );
}

//QAxObject* ActiveExcel::documentRemoveSheet(QAxObject* sheet){
// ActiveWindow.SelectedSheets.Delete
//}

QAxObject* ActiveExcel::documentClose(QAxObject* document){
  document->dynamicCall("Close(wdDoNotSaveChanges)");
  delete document;

}

void ActiveExcel::documentCloseAndSave(QAxObject *document, QVariant path){
      document -> dynamicCall("SaveAs(const QVariant&)", path);
      document->dynamicCall("Close(wdDoNotSaveChanges)");
      delete document;
}

void ActiveExcel::sheetCellPaste(QAxObject* sheet, QVariant string, QVariant row, QVariant col ){
  QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row , col);
  cell->setProperty("Value", string);
  delete cell;

}
