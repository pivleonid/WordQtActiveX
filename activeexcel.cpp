#include "activeexcel.h"

ActiveExcel::ActiveExcel()
{
  excelApplication_ = new QAxObject( "Excel.Application");
  excelApplication_->setProperty("DisplayAlerts", false);
  excelApplication_->setProperty("Visible", true);
  worcbooks_ = excelApplication_->querySubObject( "Workbooks" );
  sheets_ = new QAxObject;


}
ActiveExcel::~ActiveExcel(){
  excelApplication_->dynamicCall("Quit()");
  delete sheets_;
  delete worcbooks_;
  delete excelApplication_;
}


QAxObject* ActiveExcel::documentOpen(QVariant path){
    if (path == ""){
     return worcbooks_->querySubObject("Add");
    }
  return worcbooks_->querySubObject("Add(const QVariant &)", path);

}

QAxObject* ActiveExcel::documentAddSheet(QAxObject* worcbooks, QVariant sheet ){

        sheets_ = worcbooks->querySubObject("Sheets");
       return    sheets_->querySubObject("Add");

}

QAxObject* ActiveExcel::documentSheetActive( QVariant sheet){
    return sheets_->querySubObject( "Item(const QVariant&)", sheet );

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
  QAxObject *rangep = StatSheet->querySubObject( "Range(const QVariant&)", rowCol);
  rangep->dynamicCall("Select()");
  if (left == true)rangep->dynamicCall("HorizontalAlignment",-4152);
  if (right == true)rangep->dynamicCall("HorizontalAlignment",-4131);
  if (center == true) rangep->dynamicCall("HorizontalAlignment",-4108);
}

void ActiveExcel::sheetCellVerticalAlignment(QAxObject* sheet, QVariant rowCol, bool up, bool down, bool center){
  QAxObject *rangep = StatSheet->querySubObject( "Range(const QVariant&)", rowCol);
   rangep->dynamicCall("Select()");
   if (up == true)rangep->dynamicCall("VerticalAlignment",-4160);
   if (down == true)rangep->dynamicCall("VerticalAlignment",-4107);
   if (center == true) rangep->dynamicCall("VerticalAlignment",-4108);

}
