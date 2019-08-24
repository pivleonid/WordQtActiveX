#include "activeexcel.h"

ActiveExcel::ActiveExcel()
{
    m_flagConnect = false;
    m_flagWorkBooks = false;
    m_excelApplication = new QAxObject( "Excel.Application");
    if(m_excelApplication != nullptr)
        m_flagConnect = true;
    m_excelApplication->setProperty("DisplayAlerts", false);
    m_excelApplication->setProperty("Visible", false);
    m_workbooks = m_excelApplication->querySubObject( "Workbooks" );
    if(m_workbooks != nullptr)
        m_flagWorkBooks = true;
    m_flagClose = false;
}
//---------------------------------------------------------------------------------
ActiveExcel::~ActiveExcel(){
    if( m_flagClose == false ) //ежели приложение не было закрыто
        m_excelApplication->dynamicCall("Quit()");

    delete m_workbooks;
    delete m_excelApplication;
}
//---------------------------------------------------------------------------------
QAxObject* ActiveExcel::workbookOpen(QVariant path)
{
    QAxObject *document;
    if (path == "") document = m_workbooks->querySubObject("Add");
    else document = m_workbooks->querySubObject("Add(const QVariant &)", path);
    if(document != nullptr)
    {
        m_workSheet = document->querySubObject("Worksheets");
        if(m_workSheet == nullptr)
            return nullptr;
        m_sheets = document->querySubObject( "Sheets" );

        if(m_sheets == nullptr)
            return nullptr;
    }
    return document;
}
//---------------------------------------------------------------------------------
QStringList ActiveExcel::sheetsList()
{
    int numb = m_workSheet->dynamicCall("Count").toInt();
    QStringList names;
    for(int i = 1; i < numb+1; i++)
    {
        QAxObject* sheet = m_workSheet->querySubObject("Item(const QVariant &)", QVariant(i));
        QVariant name = sheet->dynamicCall("Name");
        names << name.toString();
        delete sheet;
    }
    return names;
}
//---------------------------------------------------------------------------------
QAxObject* ActiveExcel::workbookAddSheet(QVariant sheetName )
{
    QAxObject *active;
    active =  m_sheets->querySubObject("Add");
    active->setProperty("Name", sheetName);
    return active;
}
//---------------------------------------------------------------------------------
QAxObject* ActiveExcel::workbookSheetActive( QString sheet)
{
    return m_sheets->querySubObject( "Item(const QVariant&)", sheet );
}
//---------------------------------------------------------------------------------
bool ActiveExcel::workbookClose(QAxObject* workBook)
{
    m_flagClose = true;
    bool ret = workBook->dynamicCall("Close(wdDoNotSaveChanges)").toBool();
    delete workBook;
    return ret;
}
//---------------------------------------------------------------------------------
bool ActiveExcel::workbookCloseAndSave(QAxObject *document, QVariant path)
{
    bool ret = document -> dynamicCall("SaveAs(const QVariant&)", path).toBool();
    document->dynamicCall("Close(wdDoNotSaveChanges)");
    m_flagClose = true;
    delete document;
    return ret;
}
//---------------------------------------------------------------------------------
bool ActiveExcel::sheetCellPaste(QAxObject* sheet, QVariant string, QVariant row, QVariant col )
{
    QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row , col);
    bool ret = cell->setProperty("Value", string);
    delete cell;
    return ret;
}
//---------------------------------------------------------------------------------
bool ActiveExcel::sheetCellInsert(QAxObject* sheet, QVariant& data, QVariant row, QVariant col)
{
    QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row , col);
    if( cell == nullptr)
        return false;
    data.clear();
    data = cell->property("Value");
    delete cell;
    return true;
}
//---------------------------------------------------------------------------------
bool ActiveExcel::sheetCopyToBuf(QAxObject* sheet, QVariant rowCol)
{
    QAxObject* range = sheet->querySubObject( "Range(const QVariant&)", rowCol);
    range->dynamicCall("Select()");
    bool ret = range->dynamicCall("Copy()").toBool();
    delete range;
    return ret;
}
//---------------------------------------------------------------------------------
bool ActiveExcel::sheetPastFromBuf(QAxObject* sheet, QVariant rowCol)
{
    QAxObject* rangec = sheet->querySubObject( "Range(const QVariant&)",rowCol);
    rangec->dynamicCall("Select()");
    bool ret = rangec->dynamicCall("PasteSpecial()").toBool();
    delete rangec;
    return ret;
}
//---------------------------------------------------------------------------------
bool ActiveExcel::sheetCellMerge(QAxObject* sheet, QVariant rowCol)
{
    QAxObject* range = sheet->querySubObject( "Range(const QVariant&)", rowCol);
    range->dynamicCall("Select()");
    // устанавливаю свойство объединения.
    bool ret = range->dynamicCall("Merge()").toBool();
    delete range;
    return ret;
}
//---------------------------------------------------------------------------------
void ActiveExcel::sheetCellHeightWidth(QAxObject *sheet, QVariant RowHeight, QVariant ColumnWidth, QVariant rowCol)
{
    QAxObject *rangec = sheet->querySubObject( "Range(const QVariant&)",rowCol);
    QAxObject *razmer = rangec->querySubObject("Rows");
    razmer->setProperty("RowHeight",RowHeight);
    razmer = rangec->querySubObject("Columns");
    razmer->setProperty("ColumnWidth",ColumnWidth);
    delete razmer;
    delete  rangec;
}
//---------------------------------------------------------------------------------
void ActiveExcel::sheetCellHorizontalAlignment(QAxObject* sheet, QVariant rowCol, bool left, bool right, bool center)
{
    QAxObject *rangep = sheet->querySubObject( "Range(const QVariant&)", rowCol);
    rangep->dynamicCall("Select()");
    if (left == true)rangep->dynamicCall("HorizontalAlignment",-4152);
    if (right == true)rangep->dynamicCall("HorizontalAlignment",-4131);
    if (center == true) rangep->dynamicCall("HorizontalAlignment",-4108);
    delete rangep;
}
//---------------------------------------------------------------------------------
void ActiveExcel::sheetCellVerticalAlignment(QAxObject* sheet, QVariant rowCol, bool up, bool down, bool center)
{
    QAxObject *rangep = sheet->querySubObject( "Range(const QVariant&)", rowCol);
    rangep->dynamicCall("Select()");
    if (up == true)rangep->dynamicCall("VerticalAlignment",-4160);
    if (down == true)rangep->dynamicCall("VerticalAlignment",-4107);
    if (center == true) rangep->dynamicCall("VerticalAlignment",-4108);
    delete rangep;

}
//---------------------------------------------------------------------------------
QVariant ActiveExcel::sheetName()
{
    QAxObject* active = m_excelApplication->querySubObject("ActiveSheet");
    QVariant name = active->dynamicCall("Name");
    delete active;
    return name;
}
//---------------------------------------------------------------------------------
int ActiveExcel::sheetCellColorInsert(QAxObject* sheet, QVariant& data, QVariant row, QVariant col)
{
    QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row , col);
    if(cell == nullptr)
        return -1;
    QAxObject* interior = cell->querySubObject("Interior");
    if( interior == nullptr)
        return -2;
    data = interior->property("Color");
    delete interior;
    delete cell;
    return 0;
}
//---------------------------------------------------------------------------------
bool ActiveExcel::excelConnect()
{
    return (m_flagConnect & m_flagWorkBooks);
}
//---------------------------------------------------------------------------------
void ActiveExcel::setVisible(bool property)
{
    m_excelApplication->setProperty("Visible", property);
}





