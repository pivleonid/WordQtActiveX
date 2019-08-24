/*==================================================================*/
/*!
\brief Класс для работы с word'овскими документами при помощи ActiveQt.
\warning Добавить в .pro файл проекта QT += axcontainer
Ежели функция возвращает указатель на NULL, значит не корректная работа.
\warning При создании/ открытии документа его надо сохранить. Новому
документу автоматически присваивается индекс = 1; Позиция индекса откры-
тых ранее документов сдвигается на 1. Не сохраненный документ называется
"Документ[n]", где n = 1 до первого открытого документа.

Нумерация таблиц начинается с 1.

Совет: внимательно следите за создаваемыми объектами и указателями,
 возвращаемыми методом querySubObject(). Они не удаляются автоматически,
 нужно вызывать delete вручную. В противном случае будет эффективно
расходоваться память, и после нескольких тысяч вызовов метода
querySubObject() ваша программа и ворд в сумме займут всю память,
 но это полбеды - обращение к одной ячейке будет занимать секунду.
Это всего лишь предостережение...


В .pro файле добавить строчку QT += axcontainer
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
  bool flagWordApp;
  bool flagdoc;
  //Внутренняя функция.
  int selectionFind( QString oldString = "", QString newString = ""   /*!< [in] Старая строкаи строка для замены   */
      ,bool searchReg     = false                      /*!< [in] Учитывать регистр   */
      ,bool searchAllWord = false                      /*!< [in] Поиск целого слова  */
      ,bool searchForward = true                       /*!< [in] поиск вперед   */
      ,bool searchFormat  = true                       /*!< [in] Применить форматирование   */
      ,bool clearFormatting = true                     /*!< [in] Очистка предыдущего форматирования   */
      ,int replace = 2  );                             /*!< [in] 0- без замен, 1 = замена первого вхождения, 2 -замена всего   */


public:
  /*==================================================================*/
  /*!  \brief
  Открывает Word, делает его видимым.
  */
  ActiveWord();
  /*==================================================================*/
  /*!  \brief
  Проверка отрытия word
  */
  bool wordConnect(){
    return (flagdoc & flagWordApp);
  }

  /*==================================================================*/
  ~ActiveWord();

  void closeWordApp();
  void setVisible();
  /*==================================================================*/
  /*!  \brief
   Выбор активного документа
  */
  bool documentActive(QAxObject* document);
  /*==================================================================*/
  /*!  \brief
   Открыть документ
  \param [in] Если по умолчанию, ворд только создаст документ
  D:/var/var.docx
  \return документ.
  */

  QAxObject* documentOpen( QVariant path = "" );
  /*==================================================================*/
  /*!  \brief
   Закрытие документа без возможности его сохранения
  \param [in] document - открытый документ
  */
  bool documentClose(QAxObject* document);
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
  bool documentSave( QAxObject *document, QString path , QString fileName, QString fileFormat);
  /*==================================================================*/
  /*!  \brief
  Поиск и замена метки
  \param [in] all =  true замена всех меток. false - замена одной
  \return bool успех или неудача
  */
  bool findReplaseLabel(QString oldString, QString newString, bool all);
  /*==================================================================*/
  /*!  \brief
  Функция открывает доступ к колонтитулам, но не работает !?
  \param [in] all =  true замена всех меток. false - замена одной
  \return bool успех или неудача
  */
  bool findReplaseLabelInColontituls(QString oldString, QString newString, bool all);
  /*==================================================================*/
  /*!  \brief
  Замена меток в колонтитулах
  \param [in] doc = указатель на созданный документ
  \param [in] firstPage  true (замена только в на первой странице), false - на всех остальных
  \return bool успех или неудача
  */
  int colontitulReplaseLabel( QAxObject* doc, QString oldString, QString newString, bool firstPage);


  //----------------------------------------------------------
  /*! \brief Операции с выделенной областью*/
  //----------------------------------------------------------


  /*==================================================================*/
  /*!  \brief
   Вставка всего текста из первого документа в метку второго документа
  \return true- метка есть и замена произведена, false метка в исходном
  документе не обнаружена
  */
  bool selectionPasteText(QVariant string);
  /*==================================================================*/
  /*!  \brief
   Вставка всего текста из первого документа в метку второго документа
  \return true- метка есть и замена произведена, false метка в исходном
  документе не обнаружена
  */
  bool selectionFindAndPasteBuffer( QAxObject* document1,/*!< [in] Документ 1  */
                                  QAxObject* document2,/*!< [in] Документ 2  */
                                  QString findLabel    /*!< [in] Метка для вставки  */
                                  );

  /*==================================================================*/
  /*!  \brief
   Замена всех меток или только первой
  \return метку
  */
bool selectionFindReplaseAll(QString oldString, QString newString,
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
QVariant selectionFindColor(QString string, QVariant color, bool allText);
  /*==================================================================*/
  /*!  \brief
  Замена размера шрифта
  \param [in] string - метка;
  \param [in] fontSize - размер шрифта
  \param [in] allText - замена всех меток
  \return тип selection
  */
QVariant selectionFindSize(QString string, QVariant fontSize, bool allText);
  /*==================================================================*/
  /*!  \brief
  Замена типа шрифтa: Жирный,курсив, подчеркнутый + замена темы
  \param [in] string - метка;
  \param [in] allText - замена всех меток
  \param [in] FontName - "Times New Roman" по умолчанию
  \return тип selection
  */
int selectionFindFontname(QString string,  bool allText,bool bold = false,
                               bool italic = false, bool underline = false, QString fontName = "Arial" );

int selectionAlign( QString string, bool left, bool right, bool center );
  /*==================================================================*/
  /*!  \brief
  Выделение всего текста с возмождностью копирования в буфер
  \param [in] buffer - false- выделяю весь текст. true- и копирую его в буфер
  \return тип selection
  */
void selectionCopyAllText(bool buffer);
  /*==================================================================*/
  /*!  \brief
 Вставка теста из буфера
  \param [in] wordSelection - выделенный текст
  \return тип selection
  */
  bool selectionPasteTextFromBuffer(void);// выделенный текст
  void selectionPasteTextFromBuffer(QString findLabel);// вставка в метку
//----------------------------------------------------------
/*! \brief Операции c таблицами*/
//----------------------------------------------------------
  /*==================================================================*/
  /*!  \brief
  Вставка текста и преобразование его в таблицу
  */
  QVariant tablePaste(QList<QStringList> table, QVariant separator);
  /*==================================================================*/
  /*!  \brief

  */
  int tableGetLabels(int tableIndex, /*!< [in] индекс таблицы  */
                             int tabRow,      /*!< [in] номер шаблонный строки в таблице  */
                     QStringList&  lable/*!< [out]  Возвращает список меток в таблице. */
                             );
  /*==================================================================*/
  /*!  \brief
  Возвращает количество и список меток в таблице.
  */
  //QStringList tableGetLabels(int tableIndex);
  /*==================================================================*/
  /*!  \brief
  Возвращает количество и список !меток!  в таблице.
  Пример метки: [label]
  */
  int tableFill(QList<QStringList> tableDat_in,/*!< [in] Таблица для вставки */
                 QStringList tableLabel,        /*!< [in] Список всех меток  */
                 int tableIndex,                /*!< [in] индекс таблицы  */
                 int tabRow                     /*!< [in] номер шаблонный строки в таблице  */
                 );
  /*==================================================================*/
  /*!  \brief
  Добавляет "countLine" строк в таблицу "table".
  */
  void tableAddLine(QAxObject* table);

  /*==================================================================*/
  /*!  \brief
  Объединение ячеек в таблице.
  Внимание! Кол-во объединенных ячеек вправо- начинаются с 1. А вниз с 0!
  */
  int tableMergeCell(int tableIndex, QVariant label, int numberCol, int numberStr);

  QVariant tablesCount();


  //Добавление колонки с меткой
  int tableAddColumn(int indexTable,   /*!< [in] индекс таблицы */
                      int afterColumn,  /*!< [in] после какой колонки надо вставить колонку? */
                      QString text,     /*!< [in] название новой колонки */
                      QString label,    /*!< [in] метка колонки */
                      int row          /*!< [in] номер строки для вствки колонки*/
                                    );

  //Добавляет строку и string в таблицу, где number - номер ячейки
  int tableAddLineWithText(int tableIndex,  int number, QString string);

  int tableAutoFitWindow(int tableIndex){
      QAxObject* act = wordApplication_->querySubObject("ActiveDocument");
      if(act == NULL)
          return -1;
      QAxObject* tables = act->querySubObject("Tables");
      if(tables == NULL)
          return -2;

      QAxObject* table = tables->querySubObject("Item(const QVariant&)", tableIndex);
      if(table == NULL)
          return -3;
      table->dynamicCall("AutoFitBehavior(wdAutoFitWindow)");
      delete table;
      delete tables;
      delete act;
  }


};

/*==================================================================*/
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
/--------------------------------------------------------------/
/*  Копирование текста из документа в метку другого документа
  ActiveWord word;
  QAxObject* doc1 = word.documentOpen(true,"D:\\template1.docx");
  QAxObject* doc2 = word.documentOpen(true,"D:\\template2.docx");
  bool return_;
  return_ = word.selectionFindAndPasteBuffer(doc2,doc1, "LABEL");*/
//------------------------------------------------------------/
/*Ежели параметров больше 8, то можно воспользоваться преобразованием
 QAxObject* table = tables->querySubObject(QString("Item(%1)").arg( tableIndex).toStdString().c_str());
 %1 первый аргумент. toStdString преобразование в строку, c_str() преобразование в массив char

*/

#endif // ACTIVEWORD_H
