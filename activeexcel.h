/*==================================================================*/
/*!
\brief Класс для работы с excel'овскими документами при помощи ActiveQt.
\warning Добавить в .pro файл проекта QT += axcontainer
Ежели функция возвращает указатель на NULL, значит не корректная работа.
\warning При создании/ открытии документа его надо сохранить. Новому
документу автоматически присваивается индекс = 1; Позиция индекса откры-
тых ранее документов сдвигается на 1. Не сохраненный документ называется
"Документ[n]", где n = 1 до первого открытого документа.

Совет: внимательно следите за создаваемыми объектами и указателями,
 возвращаемыми методом querySubObject(). Они не удаляются автоматически,
 нужно вызывать delete вручную. В противном случае будет эффективно
расходоваться память, и после нескольких тысяч вызовов метода
querySubObject() ваша программа и эксель в сумме займут всю память,
 но это полбеды - обращение к одной ячейке будет занимать секунду.
Это всего лишь предостережение...

 excel -> workbook -> sheet

Пример начала работы:
ActiveExcel excel; //открываем excel
//получаем указатель на документ- используется для закрытия
QAxObject* ex1 = excel.documentOpen(QVariant(fileName_DATA));
//получаем имя основного листа
QVariant name = excel.sheetName();
//по этому имени получаем указатель на лист
QAxObject* sheet = excel.documentSheetActive(name);
Зная этот указатель возможна дальнейшая работа с excel документом
\version 1.0
*/
/*==================================================================*/


#include "qdebug.h"
#include "qaxobject.h"
#include "QStringList"


#ifndef ACTIVEEXCEL_H
#define ACTIVEEXCEL_H


class ActiveExcel
{
  QAxObject* excelApplication_; ///< файл ворда
  QAxObject* worcbooks_;        ///< Коллекция книг
  QAxObject* sheets_;           ///< Коллекция листов
  QAxObject* workSheet_;
  QAxObject* workSheets_;
  bool flagClose;
  bool flagConnect;
  bool flagWorkBooks;
public:
  ActiveExcel();
  /*==================================================================*/
  /*!  \brief
   Проверка работы excel
  */
  bool excelConnect(){
    return (flagConnect & flagWorkBooks);
  }

  void setVisible(bool a){
      excelApplication_->setProperty("Visible", a);
  }

  ~ActiveExcel();

  /*==================================================================*/
  /*!  \brief
   Открыть документ
  */
  QAxObject* workbookOpen(QVariant path = "");      /*!< [in] path = "" открывается пустой документ   */
  /*==================================================================*/
  /*!  \brief
   Получение списка листов в документе
  */
  QStringList sheetsList();
  /*==================================================================*/
  /*!  \brief
   Возвращает указатель на созданный лист
   По умолчанию создается Лист1, Лист2 ...
  */
  QAxObject* workbookAddSheet( QVariant sheetName = "" ); /*!< [in] имя листа   */
  /*==================================================================*/
  /*!  \brief
   \param [in] sheet - имя листа. По умолчанию создается Лист1, Лист2 ...
  \return  указатель листа
  */
  QAxObject* workbookSheetActive(QString sheet);
  /*==================================================================*/
  /*!  \brief
  Закрытие документа без сохранения.
  Указатель на документ будет удален внутри функции
  */
  bool workbookClose(QAxObject* workbook);   /*!< [in] указатель на созданный документ  */
  /*==================================================================*/
  /*!  \brief
   * указатель на документ будет удален внутри функции
   \param [in] path   путь для сохранения
  \return  указатель листа
  */
  bool workbookCloseAndSave(QAxObject *workbook, QVariant path);
  /*==================================================================*/
  /*!  \brief
   Установка значения в ячейку
  */
  void sheetCellPaste(QAxObject* sheet,/*!< [in] указатель листа  */
                      QVariant string, /*!< [in] строка для вставки  */
                      QVariant row, QVariant col /*!< [in] строка и столбец ячейки  */
                      );
  /*==================================================================*/
  /*!  \brief
   Получение значения из ячейки
   \return  bool true- успех
  */
  bool sheetCellInsert(QAxObject* sheet,/*!< [in] указатель листа  */
                       QVariant& data,   /*!< [in] Данные для съёма  */
                       QVariant row, QVariant col /*!< [in] строка и столбец ячейки  */
                       );
  /*==================================================================*/
  /*!  \brief
  копирование ячеек в буфер
  диспазон ячейки записывается как A1:B13
  */

  bool sheetCopyToBuf(QAxObject* sheet,/*!< [in] указатель листа  */
                      QVariant rowCol  /*!< [in] Диапазон  */
                      );
  /*==================================================================*/
  /*!  \brief
  вставка из буфера
  */
  bool sheetPastFromBuf(QAxObject* sheet,/*!< [in] указатель листа  */
                        QVariant rowCol  /*!< [in] Диапазон  */
                        );
  /*==================================================================*/
  /*!  \brief
  Объединение ячеек
  */
   bool sheetCellMerge(QAxObject* sheet,/*!< [in] указатель листа  */
                       QVariant rowCol  /*!< [in] Диапазон  */
                       );

   /*==================================================================*/
   /*!  \brief
   Ширина строк и столбцов
   */
   void sheetCellHeightWidth(QAxObject* sheet,/*!< [in] указатель листа  */
                       QVariant RowHeight, QVariant ColumnWidth,
                       QVariant rowCol/*!< [in] Диапазон  */
                       );
   /*==================================================================*/
   /*!  \brief
   Выравнивание ячеек. один из 3 параметров равен true
   */
  void sheetCellHorizontalAlignment(QAxObject* sheet,          /*!< [in] указатель листа  */
                                    QVariant rowCol,
                                    bool left = false, bool right = false, bool center = false);
  /*==================================================================*/
  /*!  \brief
  Выравнивание ячеек. один из 3 параметров равен true
  */
  void sheetCellVerticalAlignment(QAxObject* sheet,            /*!< [in] указатель листа  */
                                  QVariant rowCol,              /*!< [in] Диапазон или номер ячейки  */
                                  bool up = false, bool down = false, bool center = false);

  /*==================================================================*/
  /*!  \brief
  \return  имя активного листа
  */
  QVariant sheetName();


  /*==================================================================*/
  /*!  \brief
  \return  имя активного листа
  */
  int sheetCellColorInsert(QAxObject* sheet, QVariant& data, QVariant row, QVariant col);
};




#endif // ACTIVEEXCEL_H


