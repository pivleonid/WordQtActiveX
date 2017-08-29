#include "activeexcel.h"

ActiveExcel::ActiveExcel()
{
  excelApplication_ = new QAxObject( "Excel.Application");
  excelApplication_->setProperty("DisplayAlerts", false);
  excelApplication_->setProperty("Visible", true);
  worcbooks_ = excelApplication_->querySubObject( "Workbooks" );



}
ActiveExcel::~ActiveExcel(){
  excelApplication_->dynamicCall("Quit()");
  delete sheets_;
  delete worcbooks_;
  delete excelApplication_;
}


QAxObject* ActiveExcel::documentOpen(QVariant path){
  QAxObject *document;
  if (path == "") document = worcbooks_->querySubObject("Add");
  else document = worcbooks_->querySubObject("Add(const QVariant &)", path);

  workSheet_ = document->querySubObject("Worksheets");
  sheets_ = document->querySubObject( "Sheets" );
 return document;
}



QAxObject* ActiveExcel::documentAddSheet(QVariant sheetName ){

    QAxObject *active;
    active =  sheets_->querySubObject("Add");
    active->setProperty("Name", sheetName);
    return active;
}

QAxObject* ActiveExcel::documentSheetActive( QVariant sheet){
  //QVariant param = sheets_->dynamicCall("Count()");
    return sheets_->querySubObject( "Item(const QVariant&)", sheet );

}

//QAxObject* ActiveExcel::documentRemoveSheet(QAxObject* sheet){
// ActiveWindow.SelectedSheets.Delete
//}

void ActiveExcel::documentClose(QAxObject* document){
  document->dynamicCall("Close(wdDoNotSaveChanges)");
  delete document;

}

void ActiveExcel::documentCloseAndSave(QAxObject *document, QVariant path){
      document -> dynamicCall("SaveAs(const QVariant&)", path);
      document->dynamicCall("Close(wdDoNotSaveChanges)");
      delete document;
}
//---------------------------------------------------------------------------------
void ActiveExcel::sheetCellPaste(QAxObject* sheet, QVariant string, QVariant row, QVariant col ){
  QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row , col);
  cell->setProperty("Value", string);
  delete cell;

}
QVariant ActiveExcel::sheetCellInsert(QAxObject* sheet,  QVariant row, QVariant col){
   QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row , col);
   QVariant result = cell->property("Value");
   delete cell;
   return result;
}
//---------------------------------------------------------------------------------

void ActiveExcel::sheetCopyToBuf(QAxObject* sheet, QVariant rowCol){
  QAxObject* range = sheet->querySubObject( "Range(const QVariant&)", rowCol);
  range->dynamicCall("Select()");
  range->dynamicCall("Copy()");
}

void ActiveExcel::sheetPastFromBuf(QAxObject* sheet, QVariant rowCol){
  QAxObject* rangec = sheet->querySubObject( "Range(const QVariant&)",rowCol);
  rangec->dynamicCall("Select()");
  rangec->dynamicCall("PasteSpecial()");
}

//---------------------------------------------------------------------------------
 void ActiveExcel::sheetCellMerge(QAxObject* sheet, QVariant rowCol){
    QAxObject* range = sheet->querySubObject( "Range(const QVariant&)", rowCol);
    range->dynamicCall("Select()");
   // устанавливаю свойство объединения.
    range->dynamicCall("Merge()");
 }

void ActiveExcel::sheetCellHeightWidth(QAxObject *sheet, QVariant RowHeight, QVariant ColumnWidth, QVariant rowCol){

   QAxObject *rangec = sheet->querySubObject( "Range(const QVariant&)",rowCol);
   QAxObject *razmer = rangec->querySubObject("Rows");
   razmer->setProperty("RowHeight",RowHeight);
   razmer = rangec->querySubObject("Columns");
   razmer->setProperty("ColumnWidth",ColumnWidth);
}

void ActiveExcel::sheetCellHorizontalAlignment(QAxObject* sheet, QVariant rowCol, bool left, bool right, bool center){
  QAxObject *rangep = sheet->querySubObject( "Range(const QVariant&)", rowCol);
  rangep->dynamicCall("Select()");
  if (left == true)rangep->dynamicCall("HorizontalAlignment",-4152);
  if (right == true)rangep->dynamicCall("HorizontalAlignment",-4131);
  if (center == true) rangep->dynamicCall("HorizontalAlignment",-4108);
}

void ActiveExcel::sheetCellVerticalAlignment(QAxObject* sheet, QVariant rowCol, bool up, bool down, bool center){
  QAxObject *rangep = sheet->querySubObject( "Range(const QVariant&)", rowCol);
   rangep->dynamicCall("Select()");
   if (up == true)rangep->dynamicCall("VerticalAlignment",-4160);
   if (down == true)rangep->dynamicCall("VerticalAlignment",-4107);
   if (center == true) rangep->dynamicCall("VerticalAlignment",-4108);

}


//


void ActiveExcel::sheetProperty(QVariant sheetName,  QAxObject *workbook){
   //Проверить, есть ли такое имя
  QAxObject* sheetToCopy = workbook->querySubObject("Worksheets(const QVariant&)", "Старый лист");
  QAxObject* newSheet = workbook->querySubObject("Worksheets(const QVariant&)", "Старый лист (2)");
   QVariant param = sheets_->dynamicCall("Count()");
   QVariant param21 = workSheet_->dynamicCall("Count()");
   QAxObject* sheetsNew = sheets_->querySubObject("Add");
    QAxObject* param1 = workSheet_->querySubObject("Add()");
 param21 = workSheet_->dynamicCall("Count()");
 QVariant param213 = sheets_->dynamicCall("Codename()");
   QAxObject *StatSheet = sheets_->querySubObject( "Item(const QVariant&)", sheetName );



   int i;
   i++;

 }
