using System;
using System.Collections.Generic;
using System.Linq;

namespace SolidWorks.Interop.sldworks
{
    public enum FeatMgrTab_e
    {
        FeatureManagerTree,
        PropertyManager,
        ConfigurationManager,
        DimXpertManager,
        DisplayManager
    }

    public static class ModelDocExtension
    {
        private static Dictionary<int, FeatMgrTab_e> GetTabsMap(IModelViewManager mdlViewMgr)
        {
            return new Dictionary<int, FeatMgrTab_e>()
            {
                { mdlViewMgr.GetFeatureManagerTreeTabIndex(), FeatMgrTab_e.FeatureManagerTree },
                { mdlViewMgr.GetPropertyManagerTabIndex(), FeatMgrTab_e.PropertyManager },
                { mdlViewMgr.GetConfigurationManagerTabIndex(), FeatMgrTab_e.ConfigurationManager },
                { mdlViewMgr.GetDimXpertManagerTabIndex(), FeatMgrTab_e.DimXpertManager },
                { mdlViewMgr.GetDisplayManagerTabIndex(), FeatMgrTab_e.DisplayManager }
            };
        }

        public static void ActivateStandardFeatureManagerTab(this IModelDoc2 model, FeatMgrTab_e tab)
        {
            var mdlViewMgr = model.ModelViewManager;

            mdlViewMgr.ActiveFeatureManagerTabIndex = GetTabsMap(mdlViewMgr).First(x => x.Value == tab).Key;
        }

        public static FeatMgrTab_e GetActiveStandardFeatureManagerTab(this IModelDoc2 model)
        {
            var mdlViewMgr = model.ModelViewManager;

            FeatMgrTab_e tab;

            if (!GetTabsMap(mdlViewMgr).TryGetValue(mdlViewMgr.ActiveFeatureManagerTabIndex, out tab))
            {
                throw new NullReferenceException("Active tab is not a standard tab");
            }

            return tab;
        }
    }
}
