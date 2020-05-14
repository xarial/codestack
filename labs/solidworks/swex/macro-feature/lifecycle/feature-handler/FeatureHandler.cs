using CodeStack.SwEx.MacroFeature;
using CodeStack.SwEx.MacroFeature.Base;
using SolidWorks.Interop.sldworks;
using System.Runtime.InteropServices;

namespace CodeStack.SwEx
{
    public class LifecycleMacroFeatureParams
    {
    }

    public class LifecycleMacroFeatureHandler : IMacroFeatureHandler
    {
        public void Init(ISldWorks app, IModelDoc2 model, IFeature feat)
        {
            //feature is created or loaded
        }
        
        public void Unload(MacroFeatureUnloadReason_e reason)
        {
            switch (reason)
            {
                case MacroFeatureUnloadReason_e.Deleted:
                    //feature is deleted
                    break;

                case MacroFeatureUnloadReason_e.ModelClosed:
                    //model is closed
                    break;
            }
        }
    }

    [ComVisible(true)]
    public class LifecycleMacroFeature : MacroFeatureEx<LifecycleMacroFeatureParams, LifecycleMacroFeatureHandler>
    {
        protected override MacroFeatureRebuildResult OnRebuild(LifecycleMacroFeatureHandler handler, LifecycleMacroFeatureParams parameters)
        {
            //TODO: access handler to extract feature specific data

            return MacroFeatureRebuildResult.FromStatus(true);
        }
    }
}
