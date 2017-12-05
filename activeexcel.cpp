#include "activeexcel.h"

ActiveExcel::ActiveExcel()
{
  flagConnect = false;
  flagWorkBooks = false;
  excelApplication_ = new QAxObject( "Excel.Application");
  if(excelApplication_ != NULL)
    flagConnect = true;
  excelApplication_->setProperty("DisplayAlerts", false);
  excelApplication_->setProperty("Visible", false);
  worcbooks_ = excelApplication_->querySubObject( "Workbooks" );
  if(worcbooks_ != NULL)
    flagWorkBooks = true;
  flagClose = false;



}



ActiveExcel::~ActiveExcel(){
  if( flagClose == false ) //ежели приложение не было закрыто
    excelApplication_->dynamicCall("Quit()");

  delete worcbooks_;
  delete excelApplication_;
}


QAxObject* ActiveExcel::workbookOpen(QVariant path){
  QAxObject *document;
  if (path == "") document = worcbooks_->querySubObject("Add");
  else document = worcbooks_->querySubObject("Add(const QVariant &)", path);
  if(document != NULL){
      workSheet_ = document->querySubObject("Worksheets");
      if(workSheet_ == NULL)
        return NULL;
      sheets_ = document->querySubObject( "Sheets" );

      if(sheets_ == NULL)
        return NULL;
    }
 return document;
}

QStringList ActiveExcel::sheetsList(){
  int numb = workSheet_->dynamicCall("Count").toInt();
  QStringList names;
  for(int i = 1; i < numb+1; i++){
      QAxObject* sheet = workSheet_->querySubObject("Item(const QVariant &)", QVariant(i));
      QVariant name = sheet->dynamicCall("Name");
      names << name.toString();
      delete sheet;
    }
  return names;
}



QAxObject* ActiveExcel::workbookAddSheet(QVariant sheetName ){

    QAxObject *active;
    active =  sheets_->querySubObject("Add");
    active->setProperty("Name", sheetName);
    return active;
}

QAxObject* ActiveExcel::workbookSheetActive( QString sheet){

    return sheets_->querySubObject( "Item(const QVariant&)", sheet );

}


bool ActiveExcel::workbookClose(QAxObject* workBook){
  flagClose = true;
  bool ret = workBook->dynamicCall("Close(wdDoNotSaveChanges)").toBool();
  delete workBook;
  return ret;

}

bool ActiveExcel::workbookCloseAndSave(QAxObject *document, QVariant path){
      bool ret = document -> dynamicCall("SaveAs(const QVariant&)", path).toBool();
      document->dynamicCall("Close(wdDoNotSaveChanges)");
      flagClose = true;
      delete document;
      return ret;
}
//---------------------------------------------------------------------------------
void ActiveExcel::sheetCellPaste(QAxObject* sheet, QVariant string, QVariant row, QVariant col ){
  QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row , col);
  bool ret = cell->setProperty("Value", string);
  delete cell;

}
bool ActiveExcel::sheetCellInsert(QAxObject* sheet, QVariant& data, QVariant row, QVariant col){
   QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row , col);
   if( cell == NULL)
     return false;
   data.clear();
   data = cell->property("Value");
   delete cell;
   return true;
}
//---------------------------------------------------------------------------------

bool ActiveExcel::sheetCopyToBuf(QAxObject* sheet, QVariant rowCol){
  QAxObject* range = sheet->querySubObject( "Range(const QVariant&)", rowCol);
  range->dynamicCall("Select()");
  bool ret = range->dynamicCall("Copy()").toBool();
  delete range;
  return ret;
}

bool ActiveExcel::sheetPastFromBuf(QAxObject* sheet, QVariant rowCol){
  QAxObject* rangec = sheet->querySubObject( "Range(const QVariant&)",rowCol);
  rangec->dynamicCall("Select()");
  bool ret = rangec->dynamicCall("PasteSpecial()").toBool();
  delete rangec;
  return ret;
}

//---------------------------------------------------------------------------------
 bool ActiveExcel::sheetCellMerge(QAxObject* sheet, QVariant rowCol){
    QAxObject* range = sheet->querySubObject( "Range(const QVariant&)", rowCol);
    range->dynamicCall("Select()");
   // устанавливаю свойство объединения.
    bool ret = range->dynamicCall("Merge()").toBool();
    delete range;
    return ret;
 }

void ActiveExcel::sheetCellHeightWidth(QAxObject *sheet, QVariant RowHeight, QVariant ColumnWidth, QVariant rowCol){

   QAxObject *rangec = sheet->querySubObject( "Range(const QVariant&)",rowCol);
   QAxObject *razmer = rangec->querySubObject("Rows");
   razmer->setProperty("RowHeight",RowHeight);
   razmer = rangec->querySubObject("Columns");
   razmer->setProperty("ColumnWidth",ColumnWidth);
   delete razmer;
   delete  rangec;
}

void ActiveExcel::sheetCellHorizontalAlignment(QAxObject* sheet, QVariant rowCol, bool left, bool right, bool center){
  QAxObject *rangep = sheet->querySubObject( "Range(const QVariant&)", rowCol);
  rangep->dynamicCall("Select()");
  if (left == true)rangep->dynamicCall("HorizontalAlignment",-4152);
  if (right == true)rangep->dynamicCall("HorizontalAlignment",-4131);
  if (center == true) rangep->dynamicCall("HorizontalAlignment",-4108);
  delete rangep;
}

void ActiveExcel::sheetCellVerticalAlignment(QAxObject* sheet, QVariant rowCol, bool up, bool down, bool center){
  QAxObject *rangep = sheet->querySubObject( "Range(const QVariant&)", rowCol);
   rangep->dynamicCall("Select()");
   if (up == true)rangep->dynamicCall("VerticalAlignment",-4160);
   if (down == true)rangep->dynamicCall("VerticalAlignment",-4107);
   if (center == true) rangep->dynamicCall("VerticalAlignment",-4108);
   delete rangep;

}

QVariant ActiveExcel::sheetName(){

   QAxObject* active = excelApplication_->querySubObject("ActiveSheet");
   QVariant name = active->dynamicCall("Name");
   delete active;
   return name;
}



int ActiveExcel::sheetCellColorInsert(QAxObject* sheet, QVariant& data, QVariant row, QVariant col){
   QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row , col);
   if(cell == NULL)
       return -1;
   QAxObject* interior = cell->querySubObject("Interior");
   if( interior == NULL)
       return -2;
   data = interior->property("Color");
   delete interior;
   delete cell;
}




