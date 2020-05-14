using CodeStack.SwEx.MacroFeature;
using CodeStack.SwEx.MacroFeature.Base;
using CodeStack.SwEx.MacroFeature.Data;
using SolidWorks.Interop.sldworks;

namespace CodeStack.SwEx
{
    //returns successful regeneration without bodies
    public class RegenerationNoResultsMacroFeature : MacroFeatureEx
    {
        protected override MacroFeatureRebuildResult OnRebuild(ISldWorks app, IModelDoc2 model, IFeature feature)
        {
            return MacroFeatureRebuildResult.FromStatus(true);
        }
    }

    // returns regeneration error
    public class RegenerationRebuildErrorMacroFeature : MacroFeatureEx
    {
        protected override MacroFeatureRebuildResult OnRebuild(ISldWorks app, IModelDoc2 model, IFeature feature)
        {
            return MacroFeatureRebuildResult.FromStatus(false, "Failed to regenerate this feature");
        }
    }

    //return body without automatically assigning ids
    public class RegenerationBodyMacroFeature : MacroFeatureEx
    {
        protected override MacroFeatureRebuildResult OnRebuild(ISldWorks app, IModelDoc2 model, IFeature feature)
        {
            //use extension methods of IModeler to create a box body
            IBody2 tempBody = app.IGetModeler().CreateBox(new Point(0, 0, 0), new Vector(1, 0, 0), 0.1, 0.1, 0.1);

            return MacroFeatureRebuildResult.FromBody(tempBody, feature.GetDefinition() as IMacroFeatureData, false); 
        }
    }

    //return pattern of bodies and automatically assign entity ids
    public class RegenerationArrayOfBodiesMacroFeature : MacroFeatureEx
    {
        protected override MacroFeatureRebuildResult OnRebuild(ISldWorks app, IModelDoc2 model, IFeature feature)
        {
            IBody2[] tempBodies = null; //TODO: create temp bodies
            return MacroFeatureRebuildResult.FromBodies(tempBodies, feature.GetDefinition() as IMacroFeatureData, true);
        }
    }
}
