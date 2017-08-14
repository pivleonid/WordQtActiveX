#include <activeword.h>
#include <qwidget.h>
#include <windows.h>
//----------------------------------------------------------
ActiveWord::ActiveWord(){
  wordApplication_ =  new QAxObject("Word.Application");
  Sleep(1000);
  wordApplication_->setProperty("DisplayAlerts", false);
  Sleep(1000);
  wordApplication_->setProperty("Visible", true);
  Sleep(1000);
  documents_ = wordApplication_->querySubObject("Documents");

}
//----------------------------------------------------------
ActiveWord::~ActiveWord(){
  wordApplication_->dynamicCall("Quit()");
  delete documents_;
  delete wordApplication_;
}
//----------------------------------------------------------
void ActiveWord::documentActive(QAxObject *document){
  document->dynamicCall("Activate()");
}
//----------------------------------------------------------
QAxObject* ActiveWord::documentOpen(bool template_){
  if (!template_)
    return documents_->querySubObject("Add()");
  return  documents_->querySubObject("Add(D:\\testdot.dot)");
}
//----------------------------------------------------------
QAxObject* ActiveWord::documentOpen(bool template_, QVariant path){
  if (!template_)
    return documents_->querySubObject("Add()");
  return  documents_->querySubObject("Add(const QVariant &)", path);
}
//----------------------------------------------------------
void ActiveWord::selectionPasteText(QVariant string){
  QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
  wordSelection->dynamicCall("TypeText(const QVariant&)", string);
}
//----------------------------------------------------------
bool ActiveWord::selectionFind( QString oldString , QString newString
                         ,bool searchReg, bool searchAllWord, bool searchForward
                         , bool searchFormat, bool clearFormatting, int replace ){

    QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
    QAxObject* findString =  wordSelection->querySubObject("Find");
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

    return param.toBool();
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
  if(allText)
    return selectionFind( string, string,false,false,true,true, true, 2 );
  return selectionFind( string, string,false,false,true,true, true, 1 );
}
//----------------------------------------------------------
QVariant ActiveWord:: selectionFindFontname(QString string,  bool allText, bool bold,
                                              bool italic , bool underline, QString fontName )
{
  QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
  QAxObject* findString =  wordSelection->querySubObject("Find"); // заменить в одну строчку
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
  if(allText)
    return selectionFind( string, string,false,false,true,true, true, 2 );
  return  selectionFind( string, string,false,false,true,true, true, 1 );
}
//-----------------Возвращает указатель на объект типа selection
void ActiveWord:: selectionCopyAllText( bool buffer){
    QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
    wordSelection->dynamicCall("WholeStory()");//выделение всего
    if(buffer)
      wordSelection->dynamicCall("Copy()");//копирование выделенного в буфер обмена
}

//------------------Вставка текста из буфера
void ActiveWord:: selectionPasteTextFromBuffer(){
  QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
  wordSelection->dynamicCall("Paste()");
}
//----------------------------------------------------------
void ActiveWord::documentClose(QAxObject* document){
        document->dynamicCall("Close(wdDoNotSaveChanges)");
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
          return true;
      }

    }
    return false;
}

//----------------------------------------------------------
void ActiveWord::documentSave(QAxObject *document, QString fileName,
                                   QString fileFormat, QString path)
{
    QString all = path + fileName + fileFormat;
    QVariant param(all);
    document -> dynamicCall("SaveAs2(const QVariant&)", param);
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
    return param;
}

QStringList ActiveWord::tableGetLabels(int tableIndex, int tabRow ){
   QAxObject* act = wordApplication_->querySubObject("ActiveDocument");
   QAxObject* tables = act->querySubObject("Tables");
   //индекс указывает на искомую таблицу
   QAxObject* table = tables->querySubObject("Item(const QVariant&)", tableIndex);
   int tabColumns = table->querySubObject("Columns")->dynamicCall("count").toInt();
  // QVariant tabRow = table->querySubObject("Rows")->dynamicCall("count");//.toInt();
   QStringList lable;
   for(int i = 1; i <= tabColumns; i++){
       QAxObject* cell = table->querySubObject("Cell(const QVariant& , const QVariant&)",tabRow, i );
       QVariant str_v = cell->querySubObject("Range")->dynamicCall("Text");
       QString str = str_v.toString();
       int index = str.indexOf("]", 0 );
       str = str.mid(0, index+1);
       lable << str;
     }
return lable;
}

void ActiveWord::tableAddLine(QAxObject* table, int countLine){

  for (int i = 0; i < countLine; i++)
    table->querySubObject("Rows")->dynamicCall("Add()");

}

void ActiveWord::tableFill(QList<QStringList> tableDat_in, QStringList tableLabel, int tableIndex, int start){

  QAxObject* act = wordApplication_->querySubObject("ActiveDocument");
  QAxObject* tables = act->querySubObject("Tables");
  //список меток из шаблонной таблицы
  QStringList templateTableLabel = tableGetLabels(tableIndex, start);

  int tabColumns = templateTableLabel.count();//столбцы
  QList<int> containerIndex;
  for(int i = 0; i < tabColumns; i++)
    //во всех метках tableLabel ищу нужный индекс в стринглисте меток из шаблона
    //containerIndex.append(templateTableLabel.indexOf(tableLabel[i]));
    containerIndex.append(tableLabel.indexOf(templateTableLabel[i]));
  QAxObject* table = tables->querySubObject("Item(const QVariant&)", tableIndex);
  const int count = tableDat_in.count(); //строчки
  QAxObject* cell;
  for(int i = 1; i <= count; i++){
      if(i != 1 + start){
          if(i == count+1) return;
          tableAddLine(table, 1);//добавляю строчку
        }
      for(int j = 1; j <= tabColumns; j++){

          if(containerIndex[j-1] == -1) continue;
          cell = table->querySubObject("Cell(const QVariant& , const QVariant&)",i + start-1 , j);
          cell->querySubObject("Range")->dynamicCall("Select()");
          wordApplication_->querySubObject("Selection")->dynamicCall("TypeText(Text)", tableDat_in[i-1][containerIndex[j-1]]);

        }
    }

}


//QAxObject* table = tables->querySubObject("Item(const QVariant&)", tableIndex);
//int tabColumns = table->querySubObject("Columns")->dynamicCall("count").toInt();
//QVariant tabRow = table->querySubObject("Rows")->dynamicCall("count");//.toInt();
//QAxObject* cell = table->querySubObject("Cell(const QVariant& , const QVariant&)",4, 3);
//cell->querySubObject("Range")->dynamicCall("InsertAfter(Text)", "Это ячейка 1:1");//, "AbraCadabra");
