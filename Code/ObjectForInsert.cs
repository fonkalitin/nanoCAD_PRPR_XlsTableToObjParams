using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Multicad.Geometry;
using Multicad;
using Multicad.Objects;
using Multicad.DatabaseServices;
using Multicad.Symbols;

namespace NCadCustom.Code
{
    internal class ObjectForInsert
    {
        /// <summary>
        /// Парсинг строки эксель в объект на чертежe
        /// Завязано на номера столбцов
        /// </summary>
        /// <param name="sst"></param>
        /// <param name="headers"> названия шапок из эксель</param>
        /// <param name="objectData">строка эксель</param>
        internal ObjectForInsert(SharedStringTable sst, List<string> headers, Row objectData)
        {
            CreateObject(sst, headers, objectData);
        }
        internal Dictionary<string, string> objParams { get; set; } // Словарем определен тип данных пары значений - Строка, Строка
        private void CreateObject(SharedStringTable sst, List<string> headers, Row objectData)
        {
            IEnumerable<Cell> cells = objectData.Elements<Cell>();
            objParams = new Dictionary<string, string>();
          
            // начинаются параметры для объектов БД
            for (int iCol = 0; iCol < cells.Count(); iCol++) // Считывание параметров начинается с первого столбца таблицы iCol = 0
            {
                // Проверка типа данных в ячейке
                if (cells.ElementAt(iCol).DataType != null) // Cell cell = cells.ElementAt(iCol);
                {
                    if (cells.ElementAt(iCol).DataType == CellValues.SharedString)
                    {
                        int ssid = int.Parse(cells.ElementAt(iCol).CellValue.Text);
                        string cellValue = sst.ChildElements[ssid].InnerText.ToString(); // Извлечение текстового значения ячейки тблицы
                        objParams.Add(headers[iCol], cellValue);
                    }

                    else if (cells.ElementAt(iCol).DataType == CellValues.Number)
                    {
                        string stringVal = cells.ElementAt(iCol).CellValue.Text;
                        objParams.Add(headers[iCol], stringVal);
                    }

                    else {
                        cells.ElementAt(iCol).DataType = CellValues.Number;
                        string stringVal = cells.ElementAt(iCol).CellValue.Text;
                        objParams.Add(headers[iCol], stringVal);
                    }

                }

            }
        }

        /// <summary>
        /// Выбор объекта на чертеже и запись в него параметров из таблицы
        /// </summary>
        internal void writeParamsToObj(McDbEntity McEntityObj)
        {
            List<ExValue> paramsToChange = new List<ExValue>();
            
            foreach (KeyValuePair<string, string> paramsPair in objParams)
                    {
                     paramsToChange.Add(new ExValue(paramsPair.Key, paramsPair.Value)); // Формирование списка - пары строковых значений "ИмяПараметра-ЗначениеПараметра"
                     //McEntityObj.ObjectProperties[paramsPair.Key] = paramsPair.Value; // Непосредственно операция вставки в объект пары "Имя параметра"+""

                    //string objTAGID = McEntityObj.ObjectProperties["TAGID"] as string;
                    //string tableTAGID = paramsToChange[0].Value.ToString(); // порядковый номер из таблицы (первый столбец таблицы (TAGID))

                    //if (tableTAGID == objTAGID) // Проверка по номеру идентификатора TAGID - порядковый номер строки
                    //{
                        McEntityObj.ObjectProperties[paramsPair.Key] = paramsPair.Value;
                    //}


            }





        }

    }
}
