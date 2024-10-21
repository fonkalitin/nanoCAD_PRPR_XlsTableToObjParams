using System;
using System.Collections.Generic;
using HostMgd.ApplicationServices;
using Multicad;
using Multicad.Objects;
using Multicad.Symbols;
using Multicad.Runtime;
using Multicad.DatabaseServices;
using DocumentFormat.OpenXml.Spreadsheet;
using NCadCustom.Code;

using App = HostMgd.ApplicationServices;
using Db = Teigha.DatabaseServices;
using Ed = HostMgd.EditorInput;
using System.Runtime.Intrinsics.Arm;
using Rtm = Teigha.Runtime;
using System.Windows.Forms;
using Multicad.Symbols.Tables;

namespace NCadCustom
{
    public class Commands : IExtensionApplication
    {
        public void Initialize()
        {
            App.DocumentCollection dm = App.Application.DocumentManager;
            Ed.Editor ed = dm.MdiActiveDocument.Editor;
            string msg = "PRPR_objxldata - импорт данных из файла внешней таблицы Params.xlsx";
            ed.WriteMessage(msg);
        }
        public void Terminate()
        {
        }

        static readonly string[] excelExtentions = { ".xlsx", ".xls", ".xlsb", ".xlsm" };

        /// <summary>
        /// Импорт в объекты СПДС на чертеже параметров из таблицы эксель
        /// Параметры объектов находятся в эксель
        /// </summary>
        /// 

        
        [Rtm.CommandMethod("PRPR_objxldata", Rtm.CommandFlags.Session)]
        public static void MainCreateObjBySpreadSheet()
        {
            
            HostMgd.EditorInput.Editor ed = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            App.Document doc = App.Application.DocumentManager.MdiActiveDocument;
            //InputJig jig = new InputJig();
            //string paramsFilePath = jig.GetText("Укажите полный путь до файла параметров(Excel)", false);
            //paramsFilePath = paramsFilePath.Trim('"').ToLower();

            string dwgName = doc.Name; // метод получения полного пути и имени текущего dwg-файла
            int pos = dwgName.LastIndexOf("\\"); // позиция последнего слеша в полном пути до файла
            string dwgPath = dwgName.Remove(pos, dwgName.Length - pos); // Путь до dwg файла (без имени файла)
            string paramsFilePath = dwgPath + "\\Params.xlsx"; // путь до файла параметров (Excel)

            if (!File.Exists(paramsFilePath))
            {
                ed.WriteMessage("Выбран не существующий путь! Программа завершена.");
                return;
            }

            if (!excelExtentions.Contains(Path.GetExtension(paramsFilePath)))
            {
                ed.WriteMessage("Выбран не Excel файл! Программа завершена.");
                return;
            }

            try
            {
                ShWorker shWorker = new ShWorker(paramsFilePath);
                List<Row> dataRows = shWorker.dataRows;

                List<string> headers = shWorker.GetHeaders(shWorker.sst, dataRows.ElementAt(1));
                

                McObjectId[] idObjSelected = McObjectManager.SelectObjects("Выберите объекты: ");

                // за исключением шапки - остальное строки с данными содержат параметры объекта.
                Row rw = new Row();
                for (int iRow = 2; iRow < dataRows.Count; iRow++)
                {
                    rw = dataRows[iRow];
                    ObjectForInsert oneObject = new ObjectForInsert(shWorker.sst, headers, rw);

                        McObject SelectedObj = idObjSelected[iRow-2].GetObject(); // Перебор циклом всех выбранных объектов и построчная запись параметров

                        if (SelectedObj is McUMarker currUmarkerObject) // тип объекта - McUMarker
                        {
                            oneObject.writeParamsToObj(currUmarkerObject.DbEntity); // Вызов метода выполняющего непосредственно вставку параметров в объект
                        }

                        else if (SelectedObj is McParametricObject currParObject) // тип объекта  McParametricObject
                        {
                        oneObject.writeParamsToObj(currParObject.DbEntity); // Вызов метода выполняющего непосредственно вставку параметров в объект
                        }

                        else if (SelectedObj is McTable currTableObject) // тип объекта  таблиуа СПДС
                        {
                            oneObject.writeParamsToObj(currTableObject.DbEntity); // Вызов метода выполняющего непосредственно вставку параметров в объект
                        }
                }


                ed.WriteMessage($"Обработано позиций: {dataRows.Count - 1}");
            }
            catch (Exception e)
            {
                ed.WriteMessage($"Ошибка : {e}");
            }
        }
    }
}
