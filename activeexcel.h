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
  QAxObject* documents_;        ///< Коллекция документов
  QAxObject* sheets_;           ///< Коллекция листов
public:
  ActiveExcel();
  ~ActiveExcel();
  QAxObject* documentOpen(QVariant path = "");      /*!< [in] path = "" открывается пустой документ   */
  QAxObject* documentAddSheet( QAxObject* document, /*!< [in] документ   */
                               QVariant sheet = ""  /*!< [in] sheet имя листа   */
                             );
  //Возвращает указатель листа. По умолчанию Лист1, Лист2 ...
  QAxObject* documentSheetActive(QAxObject* sheet1 ,QVariant sheet);  /*!< [in] sheet имя листа  */
  //QAxObject* documentRemoveSheet(QAxObject* sheet);/*!< [in] sheet указатель на объект листа   */
  QAxObject* documentClose(QAxObject* document);   /*!< [in] указатель на созданный документ  */
  //путь до сохраниения и сам документ, который удалится в функции
  void documentCloseAndSave(QAxObject *document, QVariant path);  /*!< [in] путь для сохранения  */

  void sheetCellPaste(QAxObject* sheet,/*!< [in] указатель листа  */
                      QVariant string, /*!< [in] строка для вставки  */
                      QVariant row, QVariant col /*!< [in] строка и столбец ячейки  */
                      );

};

#endif // ACTIVEEXCEL_H
