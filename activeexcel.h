/*==================================================================*/
/*!
\brief Класс для работы с excel'овскими документами при помощи ActiveQt.
Ежели функция возвращает указатель на NULL, значит не корректная работа.
\warning При создании/ открытии документа его надо сохранить. Новому
документу автоматически присваивается индекс = 1; Позиция индекса откры-
тых ранее документов сдвигается на 1. Не сохраненный документ называется
"Документ[n]", где n = 1 до первого открытого документа.

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
public:
  ActiveExcel();
  ~ActiveExcel();
  QAxObject* documentOpen(QVariant path = "");      /*!< [in] path = "" открывается пустой документ   */
  //функция станавливающая коллекцию листов. Вызывать обязательно после documentOpen
  void documentGetSheet(QAxObject* document);
  //Возвращает указатель на лист
  QAxObject* documentAddSheet( QAxObject* document ); /*!< [in] документ   */

  //Возвращает указатель листа. По умолчанию Лист1, Лист2 ...
  QAxObject* documentSheetActive(QVariant sheet);  /*!< [in] sheet имя листа  */
  //QAxObject* documentRemoveSheet(QAxObject* sheet);/*!< [in] sheet указатель на объект листа   */
  QAxObject* documentClose(QAxObject* document);   /*!< [in] указатель на созданный документ  */
  //путь до сохраниения и сам документ, который удалится в функции
  void documentCloseAndSave(QAxObject *document, QVariant path);  /*!< [in] путь для сохранения  */
  //установка значения в ячейку
  void sheetCellPaste(QAxObject* sheet,/*!< [in] указатель листа  */
                      QVariant string, /*!< [in] строка для вставки  */
                      QVariant row, QVariant col /*!< [in] строка и столбец ячейки  */
                      );
  // Получение значения из ячейки
  QVariant sheetCellInsert(QAxObject* sheet,/*!< [in] указатель листа  */
                           QVariant row, QVariant col /*!< [in] строка и столбец ячейки  */
                           );
  //копирование ячеек в буфер
  // диспазон ячейки записывается как A1:B13
  void sheetCopyToBuf(QAxObject* sheet,/*!< [in] указатель листа  */
                      QVariant rowCol  /*!< [in] Диапазон  */
                      );
  //вставка из буфера
  void sheetPastFromBuf(QAxObject* sheet,/*!< [in] указатель листа  */
                        QVariant rowCol  /*!< [in] Диапазон  */
                        );

  //Объединение ячеек
   void sheetCellMerge(QAxObject* sheet,/*!< [in] указатель листа  */
                       QVariant rowCol  /*!< [in] Диапазон  */
                       );

  //Ширина строк и столбцов
   void sheetCellHeightWidth(QAxObject* sheet,/*!< [in] указатель листа  */
                       QVariant RowHeight, QVariant ColumnWidth,
                       QVariant rowCol/*!< [in] Диапазон  */
                       );
   //выравнивание ячеек. один из 3 параметров равен true
  void sheetCellHorizontalAlignment(QAxObject* sheet,          /*!< [in] указатель листа  */
                                    QVariant rowCol,
                                    bool left = false, bool right = false, bool center = false);
  void sheetCellVerticalAlignment(QAxObject* sheet,            /*!< [in] указатель листа  */
                                  QVariant rowCol,              /*!< [in] Диапазон или номер ячейки  */
                                  bool up = false, bool down = false, bool center = false);
  //void sheetCellBackgroundAndFontColor()
};

#endif // ACTIVEEXCEL_H


