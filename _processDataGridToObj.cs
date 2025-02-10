public void ProcessDataGridToObjects(DataGrid dataGrid, List<McObject> selectedObjects)
{
    // ... предыдущие проверки ...

    foreach (var item in dataGrid.Items.OfType<ParameterData>()) // Явное указание типа
    {
        if (!item.IsSelected) continue;

        McObject selectedObj = selectedObjects[dataGrid.Items.IndexOf(item)];

        // Pattern matching с объединением условий и деконструкцией
        switch (selectedObj)
        {
            case McUMarker { DbEntity: not null } currUmarker:
                ProcessUMarker(currUmarker, item);
                break;
            
            case McParametricObject { DbEntity: not null } parObj:
            case McTable { DbEntity: not null } tableObj:
                ProcessDbEntity(parObj?.DbEntity ?? tableObj.DbEntity, item);
                break;
        }
    }
}

private void ProcessUMarker(McUMarker marker, ParameterData data)
{
    ProcessDbEntity(marker.DbEntity, data);
    marker.HighLightObjects(true, Color.LightYellow); // Подсветка только для UMarker
}

private void ProcessDbEntity(McDbEntity entity, ParameterData data)
{
    // Используем рефлексию для автоматического маппинга свойств
    var properties = typeof(ParameterData).GetProperties()
        .Where(p => p.Name != nameof(ParameterData.IsSelected));

    foreach (var prop in properties)
    {
        if (entity.ObjectProperties.TryGetValue(prop.Name, out _))
        {
            entity.ObjectProperties[prop.Name] = prop.GetValue(data)?.ToString();
        }
    }
}