public void ReadDescriptionProperty()
{
    var prpMgr = App.IActiveDoc2.Extension.CustomPropertyManager[""];
    var prpName = "Description";

    string val;
    string resVal;

    if (App.IsVersionNewerOrEqual(SwVersion_e.Sw2018))
    {
        bool wasRes;
        bool linkToPrp;
        prpMgr.Get6(prpName, false, out val, out resVal, out wasRes, out linkToPrp);
    }
    else if (App.IsVersionNewerOrEqual(SwVersion_e.Sw2014))
    {
        bool wasRes;
        prpMgr.Get5(prpName, false, out val, out resVal, out wasRes);
    }
    else
    {
        prpMgr.Get4(prpName, false, out val, out resVal);
    }

    Logger.Log($"{prpName} = {resVal} [{val}]");
}