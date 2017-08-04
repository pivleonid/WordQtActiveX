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
  if (template_)
    return documents_->querySubObject("Add()");
  if(!template_)
   return  documents_->querySubObject("Add(D:\\testdot.dot)");
}
//----------------------------------------------------------
QAxObject* ActiveWord::selectionFind( QString oldString , QString newString
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
    params.operator << (QVariant("0"));// 0 =  операция поиска заканчивается, 1 = операция поиска продолжается ,
    //если достигнут начало или конец диапазона поиска
    params.operator << (QVariant(searchFormat)); //(Для применения форматирования необходимо TRUE)
    params.operator << (QVariant(newString));//Текст для замены
    params.operator << (QVariant(replace)); //2 = Замена всех; 1 = Замена первого; 0 = без замен.
    params.operator << (QVariant(false)); //облако пафоса
    params.operator << (QVariant(false)); //облако пафоса
    params.operator << (QVariant(false)); //облако пафоса
    params.operator << (QVariant(false)); //облако пафоса

    findString->dynamicCall("Execute(const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,"
                      "const QVariant&,const QVariant&,const QVariant&)",
                      params);
    return findString;
}
//----------------------------------------------------------
QAxObject* ActiveWord::selectionFindReplaseAll(QString oldString, QString newString, bool allText)
{
  if(allText)
    return  selectionFind( oldString, newString,false,false,true,false, true, 2 );
  if(!allText)
    return selectionFind( oldString, newString,false,false,true,false, true, 1 );

}

//----------------------------------------------------------
QAxObject* ActiveWord::selectionFindColor(QString string, QVariant color, bool allText){
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
    if(!allText)
     return selectionFind( string, string,false,false,true,true, true, 1 );
    }
//----------------------------------------------------------
QAxObject* ActiveWord:: selectionFindSize(QString string, QVariant fontSize, bool allText){
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
    if(!allText)
        return selectionFind( string, string,false,false,true,true, true, 1 );
}
//----------------------------------------------------------
QAxObject* ActiveWord:: selectionFindFontname(QString string,  bool allText, bool bold,
                         bool italic , bool underline, QString fontName )
{
  QAxObject* wordSelection = worgdApplication_->querySubObject("Selection");
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
    if(!allText)
        return  selectionFind( string, string,false,false,true,true, true, 1 );
}
//-----------------Возвращает указатель на объект типа selection
QAxObject* ActiveWord:: selectionCopyAllText( bool buffer){
    QAxObject* wordSelection = wordApplication_->querySubObject("Selection");
    wordSelection->dynamicCall("WholeStory()");
    if(buffer)
      wordSelection->dynamicCall("Copy()");
    return wordSelection;
}

//------------------Вставка текста из буфера
QAxObject* ActiveWord:: selectionPasteTextFromBuffer(QAxObject* wordSelection){
  wordSelection->dynamicCall("Paste()");
  return wordSelection;

}
//----------------------------------------------------------
void ActiveWord::documentClose(bool save, QAxObject* document){
    if(!save)
        document->dynamicCall("Close(wdDoNotSaveChanges)");
    if(save)
        document->dynamicCall("Close(wdSaveChanges)");
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
