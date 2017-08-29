/*==================================================================*/
/*!
\brief Класс для работы с excel'овскими документами при помощи ActiveQt.
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

\version 1.0
*/
/*==================================================================*/


#include "qdebug.h"
#include "qaxobject.h"


#ifndef ACTIVEEXCEL_H
#define ACTIVEEXCEL_H


class ActiveExcel
{
  QAxObject* excelApplication_; ///< файл ворда
  QAxObject* worcbooks_;        ///< Коллекция документов
  QAxObject* sheets_;           ///< Коллекция листов
   QAxObject* workSheet_;
   QAxObject* workSheets_;
public:
  ActiveExcel();
  ~ActiveExcel();
  /*==================================================================*/
  /*!  \brief
   Открыть документ
  */
  QAxObject* documentOpen(QVariant path = "");      /*!< [in] path = "" открывается пустой документ   */
  /*==================================================================*/
  /*!  \brief
   Возвращает указатель на созданный лист
   По умолчанию создается Лист1, Лист2 ...
  */
  QAxObject* documentAddSheet( QVariant sheetName = "" ); /*!< [in] имя листа   */
  /*==================================================================*/
  /*!  \brief
   \param [in] sheet - имя листа. По умолчанию создается Лист1, Лист2 ...
  \return  указатель листа
  */
  QAxObject* documentSheetActive(QVariant sheet);
  /*==================================================================*/
  /*!  \brief
  Закрытие документа без сохранения
  */
  void documentClose(QAxObject* document);   /*!< [in] указатель на созданный документ  */
  /*==================================================================*/
  /*!  \brief
   \param [in] path   путь для сохранения
  \return  указатель листа
  */
  void documentCloseAndSave(QAxObject *document, QVariant path);
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
  */
  QVariant sheetCellInsert(QAxObject* sheet,/*!< [in] указатель листа  */
                           QVariant row, QVariant col /*!< [in] строка и столбец ячейки  */
                           );
  /*==================================================================*/
  /*!  \brief
  копирование ячеек в буфер
  диспазон ячейки записывается как A1:B13
  */

  void sheetCopyToBuf(QAxObject* sheet,/*!< [in] указатель листа  */
                      QVariant rowCol  /*!< [in] Диапазон  */
                      );
  /*==================================================================*/
  /*!  \brief
  вставка из буфера
  */
  void sheetPastFromBuf(QAxObject* sheet,/*!< [in] указатель листа  */
                        QVariant rowCol  /*!< [in] Диапазон  */
                        );
  /*==================================================================*/
  /*!  \brief
  Объединение ячеек
  */
   void sheetCellMerge(QAxObject* sheet,/*!< [in] указатель листа  */
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

  // дополнительные функции для работы с листами
 void sheetProperty(QVariant sheetName, QAxObject *workbook);

};




#endif // ACTIVEEXCEL_H


