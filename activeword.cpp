#include <activeword.h>
#include <qwidget.h>
#include <windows.h>
//----------------------------------------------------------
ActiveWord::ActiveWord(){
  flagWordApp = false;
  flagdoc = false;
  wordApplication_ =  new QAxObject("Word.Application");
  if(wordApplication_ != NULL)
    flagWordApp = true;
  wordApplication_->setProperty("DisplayAlerts", false);
  wordApplication_->setProperty("Visible", false);
// Sleep(1000);
  documents_ = wordApplication_->querySubObject("Documents");
  if(documents_ != NULL)
    flagdoc = true;

}
void ActiveWord::setVisible(){
  wordApplication_->setProperty("Visible", true);
}
//----------------------------------------------------------
ActiveWord::~ActiveWord(){

}
void ActiveWord::closeWordApp(){
  wordApplication_->dynamicCall("Quit()");
  delete documents_;
  delete wordApplication_;
}

//----------------------------------------------------------
bool ActiveWord::documentActive(QAxObject *document){
  return (document->dynamicCall("Activate()").toBool());
}

//----------------------------------------------------------
QAxObject* ActiveWord::documentOpen(QVariant path){
  if (path == "")
    return documents_->querySubObject("Add()");
  return  documents_->querySubObject("Add(const QVariant &)", path);
}
//----------------------------------------------------------
bool ActiveWord::selectionPasteText(QVariant string){
  QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
  bool ret = wordSelection->dynamicCall("TypeText(const QVariant&)", string).toBool();
  delete wordSelection;
  return ret;
}
//----------------------------------------------------------
int ActiveWord::selectionFind( QString oldString , QString newString
                         ,bool searchReg, bool searchAllWord, bool searchForward
                         , bool searchFormat, bool clearFormatting, int replace ){



  //
   QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
   if(wordSelection == NULL)
     return -1;
   QAxObject* findString =  wordSelection->querySubObject("Find");
   if(findString == NULL)
     return -2;
    if(clearFormatting)
      findString->dynamicCall("ClearFormatting()");
    QList<QVariant> params;//Все параметры не обязательные!
    params.operator << (QVariant(oldString)); //не обязательный параметр- можно использовать ""
    params.operator << (QVariant(searchReg)); //учитывать регистр
    params.operator << (QVariant(searchAllWord));//Найти целые слова
    params.operator << (QVariant(false));// использовать подстанровочные знаки (?)
    params.operator << (QVariant(false));//звуки
    params.operator << (QVariant(false));//все словоформы
    params.operator << (QVariant(searchForward));// вперед (поиск)
    params.operator << (QVariant("1"));// 0 =  операция поиска заканчивается, 1 = операция поиска продолжается ,
    //если достигнут начало или конец диапазона поиска
    params.operator << (QVariant(searchFormat)); //(Для применения форматирования необходимо TRUE)
    params.operator << (QVariant(newString));//Текст для замены
    params.operator << (QVariant(replace)); //2 = Замена всех; 1 = Замена первого; 0 = без замен.
    params.operator << (QVariant(true)); //облако пафоса
    params.operator << (QVariant(true)); //облако пафоса
    params.operator << (QVariant(true)); //облако пафоса
    params.operator << (QVariant(true)); //облако пафоса
    QVariant param =    findString->dynamicCall("Execute(const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,const QVariant&)",
                      params);
    delete findString;
    delete wordSelection;
    return param.toInt();
}
//----------------------------------------------------------
bool ActiveWord::selectionFindAndPasteBuffer(QAxObject *document1, QAxObject *document2, QString findLabel){
  //проверить наличие метки
  documentActive(document2);
  if (selectionFind(findLabel, findLabel, false, false, true, false, true, 0) == false)
    return false;
  documentActive(document1);
  selectionCopyAllText(true);
  documentActive(document2);
  selectionFind(findLabel, "", false, false, true, false, true, 0);
  selectionPasteTextFromBuffer();
  return true;
}

//----------------------------------------------------------
bool ActiveWord::selectionFindReplaseAll(QString oldString, QString newString, bool allText)
{
  if(allText)
    return  selectionFind( oldString, newString,false,false,true,true, false, 2 );
  return selectionFind( oldString, newString,false,false,true,true, false, 1 );

}

//----------------------------------------------------------
QVariant ActiveWord::selectionFindColor(QString string, QVariant color, bool allText){
  QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
  QAxObject* findString =  wordSelection->querySubObject("Find"); // заменить в одну строчку

  findString->dynamicCall("ClearFormatting()");
  //получаем доступ к параметрам для замены
  QAxObject* replacement = findString->querySubObject("Replacement");
  //Доступ к шрифту для замены
  QAxObject* font = replacement->querySubObject("Font()");
  font->setProperty("ColorIndex", color); //например wdBlue
  delete font;
  delete replacement;
  delete findString;
  delete wordSelection;
  if(allText)
    return selectionFind( string, string,false,false,true,true, true, 2 );
  return selectionFind( string, string,false,false,true,true, true, 1 );
}
//----------------------------------------------------------
QVariant ActiveWord:: selectionFindSize(QString string, QVariant fontSize, bool allText){
  QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
  QAxObject* findString =  wordSelection->querySubObject("Find"); // заменить в одну строчку
  findString->dynamicCall("ClearFormatting()");
  //получаем доступ к параметрам для замены
  QAxObject* replacement = findString->querySubObject("Replacement");
  //Доступ к шрифту для замены
  QAxObject* font = replacement->querySubObject("Font()");
  font->setProperty("Size", fontSize);
  delete font;
  delete replacement;
  delete findString;
  delete wordSelection;
  if(allText)
    return selectionFind( string, string,false,false,true,true, true, 2 );
  return selectionFind( string, string,false,false,true,true, true, 1 );
}
//----------------------------------------------------------
int ActiveWord:: selectionFindFontname(QString string,  bool allText, bool bold,
                                              bool italic , bool underline, QString fontName )
{
  if(allText)
   bool ret = selectionFind( string, string,false,false,true,true, true, 2 );
  bool ret = selectionFind( string, string,false,false,true,true, true, 1 );
  if (ret == false)
    return -1;

  QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
  if (wordSelection == NULL)
    return -2;
  QAxObject* findString =  wordSelection->querySubObject("Find"); // заменить в одну строчку
  if (findString == NULL)
    return -3;
  findString->dynamicCall("ClearFormatting()");
  //получаем доступ к параметрам для замены
  QAxObject* replacement = findString->querySubObject("Replacement");
  //Доступ к шрифту для замены
  QAxObject* font = replacement->querySubObject("Font()");
  font->setProperty("Bold", bold);
  font->setProperty("Italic", italic);
  if(underline)
    font->setProperty("Underline", "wdUnderlineSingle");
  if(!underline)
    font->setProperty("Underline", "wdUnderlineNone");
  font->setProperty("Name", fontName);
  delete font;
  delete replacement;
  delete findString;
  delete wordSelection;
  return 0;

}
//----------------------------------------------------------
int ActiveWord::selectionAlign(QString string, bool left, bool right, bool center){


QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
if (wordSelection == NULL)
  return -1;
QAxObject* findString =  wordSelection->querySubObject("Find");
if (findString == NULL)
  return -2;
 findString->dynamicCall("ClearFormatting()");
 QList<QVariant> params;//Все параметры не обязательные!
 params.operator << (QVariant(string)); //не обязательный параметр- можно использовать ""
 params.operator << (QVariant(true)); //учитывать регистр
 params.operator << (QVariant(true));//Найти целые слова
 params.operator << (QVariant(false));// использовать подстанровочные знаки (?)
 params.operator << (QVariant(false));//звуки
 params.operator << (QVariant(false));//все словоформы
 params.operator << (QVariant(true));// вперед (поиск)
 params.operator << (QVariant("1"));// 0 =  операция поиска заканчивается, 1 = операция поиска продолжается ,
 //если достигнут начало или конец диапазона поиска
 params.operator << (QVariant(true)); //(Для применения форматирования необходимо TRUE)
 params.operator << (QVariant(string));//Текст для замены
 params.operator << (QVariant(0)); //2 = Замена всех; 1 = Замена первого; 0 = без замен.
 params.operator << (QVariant(true)); //облако пафоса
 params.operator << (QVariant(true)); //облако пафоса
 params.operator << (QVariant(true)); //облако пафоса
 params.operator << (QVariant(true)); //облако пафоса
 QVariant param =    findString->dynamicCall("Execute(const QVariant&,const QVariant&,"
                   "const QVariant&,const QVariant&,"
                   "const QVariant&,const QVariant&,"
                   "const QVariant&,const QVariant&,"
                   "const QVariant&,const QVariant&,"
                   "const QVariant&,const QVariant&,"
                   "const QVariant&,const QVariant&,const QVariant&)",
                   params);

QAxObject* paragraph;
if (left == true){
    paragraph = wordSelection->querySubObject("ParagraphFormat");
    paragraph->setProperty("Alignment","wdAlignParagraphLeft" );

  }
if (right == true){
    paragraph = wordSelection->querySubObject("ParagraphFormat");
    paragraph->setProperty("Alignment","wdAlignParagraphRight" );
  }
if (center == true){
    paragraph = wordSelection->querySubObject("ParagraphFormat");
    paragraph->setProperty("Alignment","wdAlignParagraphCenter" );
  }
 delete findString;
 delete paragraph;
 delete wordSelection;
}

//-----------------Возвращает указатель на объект типа selection
void ActiveWord:: selectionCopyAllText( bool buffer){
    QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
    wordSelection->dynamicCall("WholeStory()");//выделение всего
    if(buffer)
      wordSelection->dynamicCall("Copy()");//копирование выделенного в буфер обмена
    delete wordSelection;

}

//------------------Вставка текста из буфера
bool ActiveWord:: selectionPasteTextFromBuffer(){
  QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
  bool ret = wordSelection->dynamicCall("Paste()").toBool();
  delete wordSelection;
  return ret;
}
//------------------Вставка текста из буфера в метку
void ActiveWord:: selectionPasteTextFromBuffer(QString findLabel){

  selectionFind(findLabel, "", false, false, true, false, true, 0);
  selectionPasteTextFromBuffer();
}
//----------------------------------------------------------
bool ActiveWord::documentClose(QAxObject* document){
        return( document->dynamicCall("Close(wdDoNotSaveChanges)").toBool());
}
//----------------------------------------------------------
void ActiveWord::documentIndexClose(QAxObject* index, bool save){
    if(!save) index->dynamicCall("Close(wdDoNotSaveChanges)");
    else index->dynamicCall("Close(wdSaveChanges)");
}
//----------------------------------------------------------
bool ActiveWord::documentCheckAndClose( QString docName, bool save){
    int countDoc = documents_->dynamicCall("count").toInt();
    QAxObject *item;
    QString name;
    for(int i = 1; i <= countDoc; i++){
      item = documents_->querySubObject("Item(const QVariant &)", i);
      name = item->dynamicCall("FullName").toString();
      if(name == docName){
          if(save) documentIndexClose(item,true);
          if(!save) documentIndexClose(item, false);
          delete item;
          return true;
      }

    }
    delete item;
    return false;
}

//----------------------------------------------------------
bool ActiveWord::documentSave(QAxObject *document, QString path, QString fileName, QString fileFormat)
{
    QString all = path + fileName + "." +fileFormat;
    QVariant param(all);
    return(document -> dynamicCall("SaveAs2(const QVariant&)", param).toBool());
}
//----------------------------------------------------------
//----------------------------------------------------------
QVariant ActiveWord::tablePaste(QList<QStringList> table, QVariant separator ){
  wordApplication_->setProperty("DefaultTableSeparator(const QVariant&)", separator);

 int numRows = table.count();
 int numColumn = table[0].count();
 for( uint i =0 ; i < numRows; i++)
   for(uint j = 0; j < numColumn; j++){
       QVariant variantTable( table[i][j] ) ;

       ActiveWord::selectionPasteText(variantTable);
     }
  //создание таблицы
  QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
   wordSelection->dynamicCall("WholeStory()");
    QList<QVariant> params;//Все параметры не обязательные!
    params.operator << (QVariant(3));//[Separator]
    params.operator << (QVariant(numRows));//[NumRows]
    params.operator << (QVariant(numColumn));//[NumColumns]
    params.operator << (QVariant(false));// [InitialColumnWidth]
    //
    params.operator << (QVariant(0));                 //[Format]
    params.operator << (QVariant(true));               //  [ApplyBorders]
    params.operator << (QVariant(false));               //[ApplyShading]
    params.operator << (QVariant(true));             //[ApplyFont]
    params.operator << (QVariant(true));         //[ApplyColor]
    //
    params.operator << (QVariant(true));       //[ApplyHeadingRows]
    params.operator << (QVariant(false));      //[ApplyLastRow]
    params.operator << (QVariant(true));       // [ApplyFirstColumn]
    //
    params.operator << (QVariant(false));                  //[ApplyLastColumn]
    params.operator << (QVariant(true));                    //[AutoFit]
    params.operator << (QVariant(1));      //[AutoFitBehavior]
    params.operator << (QVariant(1));//[DefaultTableBehavior]
    QVariant param;

    param =    wordSelection->dynamicCall("ConvertToTable(const QVariant&,const QVariant&, const QVariant&,"
                                          "const QVariant&, const QVariant&, const QVariant&, const QVariant&,"
                                          "const QVariant&, const QVariant&, const QVariant&, const QVariant&,"
                                          "const QVariant&, const QVariant&, const QVariant&, const QVariant&,"
                                          "const QVariant&)", params);
    delete wordSelection;
    return param;
}

int ActiveWord::tableGetLabels(int tableIndex, int tabRow, QStringList& lable ){
    int  m = 0;
m1:
    QAxObject* act = wordApplication_->querySubObject("ActiveDocument");
    if(act == NULL){
        m++;
        goto m1;
        if(m == 10)
            return -1;
    }
    m = 0;
m2:
    QAxObject* tables = act->querySubObject("Tables");
    if(tables == NULL){
        m++;
        goto m2;
        if(m == 10)
            return -2;
    }
   //индекс указывает на искомую таблицу
   QAxObject* table = tables->querySubObject("Item(const QVariant&)", tableIndex);
   int tabColumns = table->querySubObject("Columns")->dynamicCall("count").toInt();
  // QVariant tabRow = table->querySubObject("Rows")->dynamicCall("count");//.toInt();
   QAxObject* cell;
   for(int i = 1; i <= tabColumns; i++){
       cell = table->querySubObject("Cell(const QVariant& , const QVariant&)",tabRow, i );
       QVariant str_v = cell->querySubObject("Range")->dynamicCall("Text");
       QString str = str_v.toString();
       int index = str.indexOf("]", 0 );
       str = str.mid(0, index+1);
       if(str.isEmpty())
         continue;
       lable << str;
     }
   delete cell;
   delete table;
   delete tables;
   delete act;
return 0;
}

void ActiveWord::tableAddLine(QAxObject* table){
  QAxObject* rows;
  rows = table->querySubObject("Rows");//->dynamicCall("Add()");
  rows->dynamicCall("Add()");
  delete rows;
}

//tableLabel метки могут не совпадать с метками в шаблонном документе
int ActiveWord::tableFill(QList<QStringList> tableDat_in, QStringList tableLabel, int tableIndex, int start){

    int m = 0;
m1:
    QAxObject* act = wordApplication_->querySubObject("ActiveDocument");
    if( act == NULL){
        m++;
        if(m == 10)
            return -1;
        goto m1;
    }
    m = 0;
m2:
    QAxObject* tables = act->querySubObject("Tables");
    if( tables == NULL){
        m++;
        if(m == 10)
            return -2;
        goto m2;

    }
  //список меток из шаблонной таблицы
  QStringList templateTableLabel;
  int ret = tableGetLabels(tableIndex, start, templateTableLabel);
if(ret < 0)
    return -4;
  int tabColumns = templateTableLabel.count();//столбцы
  QList<int> containerIndex;
  for(int i = 0; i < tabColumns; i++)
    //во всех метках tableLabel ищу нужный индекс в стринглисте меток из шаблона
    //containerIndex.append(templateTableLabel.indexOf(tableLabel[i]));
    containerIndex.append(tableLabel.indexOf(templateTableLabel[i]));
  QAxObject* table = tables->querySubObject("Item(const QVariant&)", tableIndex);
  if( table == NULL)
    return -3;
  //количество добаввленных строк
  const int count = tableDat_in.count();

  for(int i = 1; i <= count; i++){
      if(i != 1 + start){
          if(i == count+1)
            return 0;
          tableAddLine(table);//добавляю строчку
        }
      for(int j = 1; j <= tabColumns; j++){
          //ежели элемент не найден в таблице меток
          if(containerIndex[j-1] == -1)
            continue;
          //if( tableDat_in[j].count() < j)
            //continue;
          //b = tableDat_in[i].count();
          QAxObject* cell = table->querySubObject("Cell(const QVariant& , const QVariant&)",i + start-1 , j);
          if( cell == NULL)
            return -4;
          cell->querySubObject("Range")->dynamicCall("Select()");
          //если метка стоит, а замещающая строка пустая - то метка останется!
          if( i == 1){ //сделать только в первом прогоне
            QAxObject* sel =wordApplication_->querySubObject("Selection");
            if( sel == NULL)
              return -5;
            sel->dynamicCall("Cut()");
            QString s = tableDat_in[i-1][containerIndex[j-1]];
            sel->dynamicCall("TypeText(Text)", QVariant(s));
            delete sel;
            delete cell;
            continue;
            }
          QAxObject* sel = wordApplication_->querySubObject("Selection");
          if( sel == NULL)
            return -6;
          sel->dynamicCall("TypeText(Text)", tableDat_in[i-1][containerIndex[j-1]]);
          delete sel;
          delete cell;
        }
    }
  delete table;
  delete tables;
  delete act;
  return 0;
}


//QAxObject* table = tables->querySubObject("Item(const QVariant&)", tableIndex);
//int tabColumns = table->querySubObject("Columns")->dynamicCall("count").toInt();
//QVariant tabRow = table->querySubObject("Rows")->dynamicCall("count");//.toInt();
//QAxObject* cell = table->querySubObject("Cell(const QVariant& , const QVariant&)",4, 3);
//cell->querySubObject("Range")->dynamicCall("InsertAfter(Text)", "Это ячейка 1:1");//, "AbraCadabra");


int ActiveWord::tableMergeCell(int tableIndex, QVariant label, int numberCol, int numberStr){

    QAxObject* act = wordApplication_->querySubObject("ActiveDocument");
    if(act == NULL)
        return -1;
    QAxObject* tables = act->querySubObject("Tables");
    if(tables == NULL)
        return -2;
    QAxObject* table = tables->querySubObject("Item(const QVariant&)", tableIndex);
    if(table == NULL)
        return -3;

    //
    QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
    if(wordSelection == NULL)
        return -4;
    QAxObject* findString =  wordSelection->querySubObject("Find");
    if(findString == NULL)
        return -5;
    findString->dynamicCall("ClearFormatting()");
    QList<QVariant> params;//Все параметры не обязательные!
    params.operator << (QVariant(label)); //не обязательный параметр- можно использовать ""
    params.operator << (QVariant(false)); //учитывать регистр
    params.operator << (QVariant(false));//Найти целые слова
    params.operator << (QVariant(false));// использовать подстанровочные знаки (?)
    params.operator << (QVariant(false));//звуки
    params.operator << (QVariant(false));//все словоформы
    params.operator << (QVariant(true));// вперед (поиск)
    params.operator << (QVariant("1"));// 0 =  операция поиска заканчивается, 1 = операция поиска продолжается ,
    //если достигнут начало или конец диапазона поиска
    params.operator << (QVariant(true)); //(Для применения форматирования необходимо TRUE)
    params.operator << (QVariant(""));//Текст для замены
    params.operator << (QVariant(0)); //2 = Замена всех; 1 = Замена первого; 0 = без замен.
    params.operator << (QVariant(true)); //облако пафоса
    params.operator << (QVariant(true)); //облако пафоса
    params.operator << (QVariant(true)); //облако пафоса
    params.operator << (QVariant(true)); //облако пафоса
    QVariant param =    findString->dynamicCall("Execute(const QVariant&,const QVariant&,"
                                                "const QVariant&,const QVariant&,"
                                                "const QVariant&,const QVariant&,"
                                                "const QVariant&,const QVariant&,"
                                                "const QVariant&,const QVariant&,"
                                                "const QVariant&,const QVariant&,"
                                                "const QVariant&,const QVariant&,const QVariant&)",
                                                params);
    //

    wordSelection->dynamicCall("SelectCell");

    wordSelection->dynamicCall("MoveRight(const QVariant&, const QVariant&, const QVariant&)", 1, numberCol, 1) ;
    wordSelection->dynamicCall("MoveDown(const QVariant&, const QVariant&, const QVariant&)", 5 , numberStr, 1 );

    QAxObject* cells =  wordSelection->querySubObject("Cells");
    if(cells == NULL)
        return -6;
    cells->dynamicCall("Merge()");

    //wordSelection->dynamicCall("Delete(wdCharacter, 1)");
    // wordSelection->dynamicCall("TypeText(const QVariant&)", label);

    delete cells;
    delete findString;
    delete wordSelection;
    delete table;
    delete tables;
    delete act;

}



QVariant ActiveWord::tablesCount(){
   QAxObject* act = wordApplication_->querySubObject("ActiveDocument");
   QAxObject* tables = act->querySubObject("Tables");
   QVariant count = tables->dynamicCall("Count");
   delete tables;
   delete act;
   return count;
}


bool ActiveWord::findReplaseLabel(QString oldString, QString newString, bool all){
  if (all == true)
    return selectionFind(  oldString,  newString, true, true, true, true, true,2);

  if(all == false)
    return selectionFind(  oldString,  newString, false, false, true, true, true,1);

}
bool ActiveWord::findReplaseLabelInColontituls(QString oldString, QString newString, bool all){
  //ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter //10
  QAxObject* activwin = wordApplication_->querySubObject("ActiveWindow");
  QAxObject* pane = activwin->querySubObject("ActivePane");
  QAxObject* view = pane->querySubObject("View");
  view->setProperty("SeekView", 10);
  view->setProperty("SeekView", 9);

  if (all == true)
   return selectionFind(  oldString,  newString, false, false, true, true, true,2);


  if(all == false)
      return selectionFind(  oldString,  newString, false, false, true, true, true,1);

}

int ActiveWord::colontitulReplaseLabel( QAxObject* doc, QString oldString, QString newString, bool firstPage){
  QAxObject* stor = doc->querySubObject("StoryRanges");
  if(stor == NULL)
    return -1;
  QAxObject*  range;
  if(firstPage == true)
    range = stor->querySubObject("Item( wdFirstPageFooterStory )" ); // последующие стр wdPrimaryFooterStory
  if(firstPage == false)
    range = stor->querySubObject("Item( wdPrimaryFooterStory )" ); // последующие стр wdPrimaryFooterStory
  if(range == NULL)
    return -2;
  QAxObject*  findString =  range->querySubObject("Find");
  if(findString == NULL)
    return -3;
  QList<QVariant> params;//Все параметры не обязательные!
  params.operator << (QVariant(oldString)); //не обязательный параметр- можно использовать ""
  params.operator << (QVariant(false)); //учитывать регистр
  params.operator << (QVariant(true));//Найти целые слова
  params.operator << (QVariant(false));// использовать подстанровочные знаки (?)
  params.operator << (QVariant(false));//звуки
  params.operator << (QVariant(false));//все словоформы
  params.operator << (QVariant(true));// вперед (поиск)
  params.operator << (QVariant("1"));// 0 =  операция поиска заканчивается, 1 = операция поиска продолжается ,
  //если достигнут начало или конец диапазона поиска
  params.operator << (QVariant(true)); //(Для применения форматирования необходимо TRUE)
  params.operator << (QVariant(newString));//Текст для замены
  params.operator << (QVariant(2)); //2 = Замена всех; 1 = Замена первого; 0 = без замен.
  params.operator << (QVariant(true)); //облако пафоса
  params.operator << (QVariant(true)); //облако пафоса
  params.operator << (QVariant(true)); //облако пафоса
  params.operator << (QVariant(true)); //облако пафоса
  findString->dynamicCall("Execute(const QVariant&,const QVariant&,"
                                              "const QVariant&,const QVariant&,"
                                              "const QVariant&,const QVariant&,"
                                              "const QVariant&,const QVariant&,"
                                              "const QVariant&,const QVariant&,"
                                              "const QVariant&,const QVariant&,"
                                              "const QVariant&,const QVariant&,const QVariant&)",
                                              params);

  delete range;
  delete stor;
  return 0;


    }


int ActiveWord::tableAddColumn(int indexTable, int afterColumn, QString text, QString label, int row){

    QAxObject* act = wordApplication_->querySubObject("ActiveDocument");
    if( act == NULL)
        return -1;
    QAxObject* tables = act->querySubObject("Tables");
    if( tables == NULL)
        return -2;
    //индекс указывает на искомую таблицу
    QAxObject* table = tables->querySubObject("Item(const QVariant&)", indexTable);
    if( table == NULL)
        return -3;
    QAxObject* columns =  table->querySubObject("Columns");
    if( columns == NULL)
        return -4;
    QAxObject* col = columns->querySubObject("Item(const QVariant&)", afterColumn);
    if( col == NULL)
        return -5;
    col->dynamicCall("Select()");
    //Selection.InsertColumnsRight
    QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
    if( wordSelection == NULL)
        return -6;
    wordSelection->dynamicCall("InsertColumnsRight()");


    //вставка названия колонки
    QAxObject* cell = table->querySubObject("Cell(const QVariant& , const QVariant&)",row ,afterColumn + 1);
    if( cell == NULL)
        return -7;
    cell->querySubObject("Range")->dynamicCall("Select()");
    QAxObject* sel =wordApplication_->querySubObject("Selection");
    if( sel == NULL)
        return -8;
    sel->dynamicCall("TypeText(Text)", QVariant(text));
    //вставка метки в ячейку
    cell = table->querySubObject("Cell(const QVariant& , const QVariant&)",row + 1 ,afterColumn + 1);
    cell->querySubObject("Range")->dynamicCall("Select()");
    sel =wordApplication_->querySubObject("Selection");
    sel->dynamicCall("TypeText(Text)", QVariant(label));

    delete sel;
    delete cell;
    delete wordSelection;
    delete col;
    delete columns;
    delete       table;
    delete       tables;
    delete       act;

}


int ActiveWord::tableAddLineWithText(int tableIndex, int number, QString string){

    QAxObject* act = wordApplication_->querySubObject("ActiveDocument");
    if(act == NULL)
        return -1;
    QAxObject* tables = act->querySubObject("Tables");
    if(tables == NULL)
        return -2;

    QAxObject* table = tables->querySubObject("Item(const QVariant&)", tableIndex);
    if(table == NULL)
        return -3;

    QAxObject* rows =  table->querySubObject("Rows");
    if(rows == NULL)
        return -4;

    rows->dynamicCall("Add()");

    int tabColumns = rows->dynamicCall("count").toInt();

    QAxObject* cell = table->querySubObject("Cell(const QVariant& , const QVariant&)",tabColumns , number);
    if(cell == NULL)
        return -5;
    cell->querySubObject("Range")->dynamicCall("Select()");
    QAxObject* sel =wordApplication_->querySubObject("Selection");
    if(sel == NULL)
        return -6;
    sel->dynamicCall("TypeText(Text)", QVariant(string));

    delete sel;
    delete cell;
    delete  rows;
    delete table;
    delete tables;
    delete act;
}

