public void GetTolerance(IDimension dim)
{
    var dimTol = dim.Tolerance;

    double maxTol;
    double minTol;

    if (App.IsVersionNewerOrEqual(SwVersion_e.Sw2015, 3))
    {
        dimTol.GetMinValue2(out minTol);
        dimTol.GetMaxValue2(out maxTol);
    }
    else
    {
        minTol = dimTol.GetMinValue();
        maxTol = dimTol.GetMaxValue();
    }
}