using CodeStack.SwEx.MacroFeature;
using CodeStack.SwEx.MacroFeature.Attributes;
using CodeStack.SwEx.MacroFeature.Base;
using CodeStack.SwEx.MacroFeature.Data;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace CodeStack.SwEx
{
    public class DimensionMacroFeatureParams
    {
        [ParameterDimension(swDimensionType_e.swLinearDimension)]
        public double FirstDimension { get; set; } = 0.01;

        [ParameterDimension(swDimensionType_e.swRadialDimension)]
        public double SecondDimension { get; set; }
    }    
}
