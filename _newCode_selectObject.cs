Row rw = new Row();
for (int iRow = 2; iRow < dataRows.Count; iRow++) // Третья строка со значениями параметров в таблице (iRow = 2)
{
    rw = dataRows[iRow];
    ObjectForInsert oneObject = new ObjectForInsert(shWorker.sst, headers, rw);
    McObject SelectedObj = idObjSelected[iRow-2].GetObject(); // Перебор выбранных объектов для построчной записи

    // Обработка UMarker с подсветкой
    if (SelectedObj is McUMarker currUmarker)
    {
        oneObject.writeParamsToObj(currUmarker.DbEntity);
        currUmarker.HighLightObjects(true, Color.LightYellow); // Подсветка связанных объектов
    }
    // Общая обработка параметрических объектов и таблиц
    else if ((SelectedObj as McParametricObject)?.DbEntity is {} parEntity 
          || (SelectedObj as McTable)?.DbEntity is {} tableEntity)
    {
        oneObject.writeParamsToObj(parEntity ?? tableEntity); // Вставка параметров
    }
}