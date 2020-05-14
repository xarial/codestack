using CodeStack.SwEx.PMPage.Attributes;
using SolidWorks.Interop.swconst;

public class NumberBoxDataModel
{

    [NumberBoxOptions(swNumberboxUnitType_e.swNumberBox_Length, 0, 1000, 0.01, true, 0.02, 0.001,
        swPropMgrPageNumberBoxStyle_e.swPropMgrPageNumberBoxStyle_Thumbwheel)]
    public double Length { get; set; }
}