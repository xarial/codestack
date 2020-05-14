private bool OnRebuild(DocumentHandler docHandler, RebuildState_e type)
{
    Logger.Log($"'{docHandler.Model.GetTitle()}' rebuilt ({type})");

    if(type == RebuildState_e.PreRebuild)
    {
        //return false to cancel regeneration
    }

    return true;
}

private void OnInitialized(DocumentHandler docHandler)
{
    Logger.Log($"'{docHandler.Model.GetTitle()}' initialized");
}

private bool OnSelection(DocumentHandler docHandler, swSelectType_e selType, SelectionState_e type)
{
    Logger.Log($"'{docHandler.Model.GetTitle()}' selection ({type}) of {selType}");

    if (type != SelectionState_e.UserPreSelect) //dynamic selection
    {
        //return false to cancel selection
    }

    return true;
}

private bool OnSave(DocumentHandler docHandler, string fileName, SaveState_e type)
{
    Logger.Log($"'{docHandler.Model.GetTitle()}' saving ({type})");

    if (type == SaveState_e.PreSave)
    {
        //return false to cancel saving
    }

    return true;
}

private void OnItemModified(DocumentHandler docHandler, ItemModificationAction_e type, swNotifyEntityType_e entType, string name, string oldName = "")
{
    Logger.Log($"'{docHandler.Model.GetTitle()}' item modified ({type}) of {entType}. Name: {name} (from {oldName})");
}

private void OnCustomPropertyModified(DocumentHandler docHandler, CustomPropertyModifyData[] modifications)
{
    foreach (var mod in modifications)
    {
        Logger.Log($"'{docHandler.Model.GetTitle()}' custom property '{mod.Name}' changed ({mod.Action}) in '{mod.Configuration}' to '{mod.Value}'");
    }
}

private void OnAccess3rdPartyData(DocumentHandler docHandler, Access3rdPartyDataState_e state)
{
    Logger.Log($"'{docHandler.Model.GetTitle()}' accessing 3rd party data ({state})");
}

private void OnConfigurationOrSheetChanged(DocumentHandler docHandler, ConfigurationChangeState_e type, string confName)
{
    Logger.Log($"'{docHandler.Model.GetTitle()}' configuration {confName} changed ({type})");
}

private void OnDimensionChange(DocumentHandler docHandler, IDisplayDimension dispDim)
{
    var dim = dispDim.GetDimension2(0);

    Logger.Log($"'{docHandler.Model.GetTitle()}' dimension change: {dim.FullName} = {dim.Value}");

    Marshal.ReleaseComObject(dim);
    dim = null;
}

private void OnActivated(DocumentHandler docHandler)
{
    Logger.Log($"'{docHandler.Model.GetTitle()}' activated");
}