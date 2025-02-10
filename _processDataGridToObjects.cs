// Основной метод для обработки данных из DataGrid и записи их в выбранные объекты
public void ProcessDataGridToObjects(DataGrid dataGrid, List<McObject> selectedObjects)
{
    // Проверка минимального количества столбцов
    if (dataGrid.Columns.Count < 2)
        throw new InvalidOperationException("DataGrid должен иметь как минимум два столбца.");

    // Валидация первого столбца как чекбокса
    var checkBoxColumn = dataGrid.Columns[0] as DataGridCheckBoxColumn
        ?? throw new InvalidOperationException("Первый столбец должен быть DataGridCheckBoxColumn.");

    // Получение привязки данных для чекбокса
    var checkBoxBinding = checkBoxColumn.Binding as Binding
        ?? throw new InvalidOperationException("Чекбокс столбец должен иметь валидную привязку.");

    string checkBoxPropertyPath = checkBoxBinding.Path.Path;

    // Обработка строк данных
    for (int iRow = 0; iRow < dataGrid.Items.Count; iRow++)
    {
        var item = dataGrid.Items[iRow];

        // Пропуск неотмеченных строк
        if (!GetCheckBoxValue(item, checkBoxPropertyPath)) continue;

        // Получаем выбранный объект для текущей строки
        McObject selectedObj = selectedObjects[iRow];

        // Обработка каждого столбца (начиная со второго, так как первый - чекбокс)
        for (int iCol = 1; iCol < dataGrid.Columns.Count; iCol++)
        {
            var column = dataGrid.Columns[iCol] as DataGridBoundColumn;
            if (column == null) continue;

            var bindingPath = (column.Binding as Binding)?.Path.Path;
            if (bindingPath == null) continue;

            // Получаем значение из DataGrid
            var value = BindingEvaluator.GetValue(item, bindingPath);
            string paramValue = value?.ToString() ?? string.Empty;

            // Записываем значение в объект
            if (selectedObj is McUMarker currUmarker)
            {
                currUmarker.DbEntity.ObjectProperties[bindingPath] = paramValue;
                currUmarker.HighLightObjects(true, Color.LightYellow); // Подсветка связанных объектов
            }
            else if (selectedObj is McParametricObject currParObject)
            {
                currParObject.DbEntity.ObjectProperties[bindingPath] = paramValue;
            }
            else if (selectedObj is McTable currTableObject)
            {
                currTableObject.DbEntity.ObjectProperties[bindingPath] = paramValue;
            }
        }
    }
}

// Вспомогательный метод для получения значения чекбокса
private bool GetCheckBoxValue(object item, string propertyPath)
{
    var propertyInfo = item.GetType().GetProperty(propertyPath);
    return propertyInfo != null && (bool)propertyInfo.GetValue(item);
}