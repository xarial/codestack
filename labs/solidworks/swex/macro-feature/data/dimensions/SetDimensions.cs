using CodeStack.SwEx.MacroFeature;
using CodeStack.SwEx.MacroFeature.Base;
using CodeStack.SwEx.MacroFeature.Data;
using SolidWorks.Interop.sldworks;

namespace CodeStack.SwEx
{
    public class MyDimMacroFeature : MacroFeatureEx<DimensionMacroFeatureParams>
    {
        protected override void OnSetDimensions(ISldWorks app, IModelDoc2 model,
            IFeature feature, MacroFeatureRebuildResult rebuildResult, DimensionDataCollection dims,
            DimensionMacroFeatureParams parameters)
        {
            dims[0].SetOrientation(new Point(0, 0, 0), new Vector(0, 1, 0));

            dims[1].SetOrientation(new Point(0, 0, 0), new Vector(0, 0, 1));
        }
    }
}
