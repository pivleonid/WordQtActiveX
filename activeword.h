/*==================================================================*/
/*!
\brief Класс для работы с word'овскими документами при помощи ActiveQt.
Ежели функция возвращает указатель на NULL, значит не корректная работа.
\warning При создании/ открытии документа его надо сохранить. Новому
документу автоматически присваивается индекс = 1; Позиция индекса откры-
тых ранее документов сдвигается на 1. Не сохраненный документ называется
"Документ[n]", где n = 1 до первого открытого документа.

\version 1.0
*/
/*==================================================================*/
#ifndef ACTIVEWORD_H
#define ACTIVEWORD_H

#include "qdebug.h"
#include "qaxobject.h"

class ActiveWord{

  QAxObject* wordApplication_; ///< файл ворда
  QAxObject* documents_;       ///< Коллекция документов

public:
  /*==================================================================*/
  /*!  \brief
  Открывает Word, делает его видимым.
  */
  ActiveWord();
  /*==================================================================*/
  ~ActiveWord();
  void documentActive(QAxObject* document);
  /*==================================================================*/
  /*!  \brief
   Открыть документ
  \param [in] template_ - true - открыть шаблон, false- создать документ
  \return документ.
  */
  QAxObject* documentOpen( bool template_ );
  /*==================================================================*/
  /*!  \brief
  документ должен быть создан или сохранен функцией documentSave(...);
  \param [in] document - открытый документ
  \param [in] save - сорхранить документ? .
  */
  void documentClose(bool save, QAxObject* document);
  /*==================================================================*/
  /*!  \brief
  документ должен быть создан или сохранен функцией documentSave(...);
  \param [in] index - индекс элемента;
  \param [in] save - сорхранить документ?
  */
  void documentIndexClose(QAxObject* index, bool save);
  /*==================================================================*/
  /*!  \brief
  документ должен быть создан или сохранен функцией documentSave(...);
  \param [in] docName - имя документа;
  \param [in] save - сорхранить документ?
  \return bool. bool == false - такого документа нет
  */
  bool documentCheckAndClose( QString docName, bool save);
  /*==================================================================*/
  /*!  \brief
  Сохранить как
  \param [in] document - документ;
  \param [in] path - путь до файла
  \param [in] fileName - имя файла
  \param [in] fileFormat - формат файла
  */
  void documentSave(QAxObject *document, QString path, QString fileName, QString fileFormat);
  //----------------------------------------------------------
  /*! \brief Операции с выделенной областью*/
  //----------------------------------------------------------
  /*==================================================================*/
  /*!  \brief
   Выбор меток
  \return метку
  */
  QAxObject* selectionFind( QString oldString = "", QString newString = ""   /*!< [in] Старая строкаи строка для замены   */
      ,bool searchReg     = false                      /*!< [in] Учитывать регистр   */
      ,bool searchAllWord = false                      /*!< [in] Поиск целого слова  */
      ,bool searchForward = true                       /*!< [in] поиск вперед   */
      ,bool searchFormat  = true                       /*!< [in] Применить форматирование   */
      ,bool clearFormatting = true                     /*!< [in] Очистка предыдущего форматирования   */
      ,int replace = 2  );                             /*!< [in] 0- без замен, 1 = замена первого вхождения, 2 -замена всего   */

  /*==================================================================*/
  /*!  \brief
   Замена всех меток или только первой
  \return метку
  */
  QAxObject* selectionFindReplaseAll(QString oldString, QString newString,
                                     bool allText  /*!< [in] Замена всех меток  */
                                     );
  //----------------------------------------------------------
  ///Набор цветов
  enum color{
    wdBlack = 1, ///< черный цвет
    wdBlue, wdBrightGreen =4, wdRed = 6,wdYellow,  wdDarkBlue = 9, wdGreen = 11, wdViolet, wdDarkRed, wdDarkYellow,

  };
  /*==================================================================*/
  /*!  \brief
  Замена цвета метки
  \param [in] string - метка;
  \param [in] color - смотри перечисление набора цветов
  \param [in] allText - замена всех меток
  \return тип selection
  */
  QAxObject* selectionFindColor(QString string, QVariant color, bool allText);
  /*==================================================================*/
  /*!  \brief
  Замена размера шрифта
  \param [in] string - метка;
  \param [in] fontSize - размер шрифта
  \param [in] allText - замена всех меток
  \return тип selection
  */
  QAxObject* selectionFindSize(QString string, QVariant fontSize, bool allText);
  /*==================================================================*/
  /*!  \brief
  Замена типа шрифтa: Жирный,курсив, подчеркнутый + замена темы
  \param [in] string - метка;
  \param [in] allText - замена всех меток
  \param [in] FontName - "Times New Roman" по умолчанию
  \return тип selection
  */
  QAxObject* selectionFindFontname(QString string,  bool allText,bool bold = false,
                                   bool italic = false, bool underline = false, QString FontName = "Times New Roman");
  /*==================================================================*/
  /*!  \brief
  Выделение всего текста с возмождностью копирования в буфер
  \param [in] buffer - false- выделяю весь текст. true- и копирую его в буфер
  \return тип selection
  */
  QAxObject* selectionCopyAllText(bool buffer);
  /*==================================================================*/
  /*!  \brief
 Вставка теста из буфера
  \param [in] wordSelection - выделенный текст
  \return тип selection
  */
  QAxObject* selectionPasteTextFromBuffer(QAxObject* wordSelection);// выделенный текст
};
  /*
/-- Пример выделения всего текста, копирования его в буфер, а также
/--  поиск метки и вставка на ее место текста из буфера
QAxObject* buf = AllTextAndCopySelection(word, true);
QAxObject* findWord;
findWord = findSelection(word, "LABEL", "", false, false, true, false, true, 0);
PasteTextFromBufferSelection(buf);
/--------------------------------------------------------------/
/--Пример работы с двумя активными документами
QAxObject *documents = word->querySubObject("Documents"); //получаем коллекцию документов
QAxObject *document = documents->querySubObject("Add(D:\\testdot.docx)"); //
QAxObject *document1 = documents->querySubObject("Add()");
///По умолчанию активный документ, это последнеоткрытый документ - document1.
QAxObject *selection1 = word->querySubObject("Selection");
selection1->dynamicCall("TypeText(Hellllllo)");Вставка текста
activeDocument(document);// активируем предыдущий документ

QAxObject *selection = word->querySubObject("Selection");
selection->dynamicCall("TypeText(Hellllllo)");

/--------------------------------------------------------------/
/--Пример сохранения документа
QAxObject *selection = word->querySubObject("Selection");
selection->dynamicCall("TypeText(adasdasdasdasd)");
saveDocument(document, "Word",".docx", "D:\\");
 */

#endif // ACTIVEWORD_H
