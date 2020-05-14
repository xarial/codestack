using CodeStack.SwEx.MacroFeature;
using SolidWorks.Interop.sldworks;

namespace CodeStack.SwEx
{
    public class EditMacroFeatureDefinitionParameters
    {
        //TODO: add properties
    }

    public class EditMacroFeatureDefinition:MacroFeatureEx<EditMacroFeatureDefinitionParameters>
    {
        protected override bool OnEditDefinition(ISldWorks app, IModelDoc2 model, IFeature feature)
        {
            var featData = feature.GetDefinition() as IMacroFeatureData;

            //rollback feature
            featData.AccessSelections(model, null);

            //read current parameters
            var parameters = GetParameters(feature, featData, model);

            var res = ShowPage(parameters);

            if (res)
            {
                //set parameters and update feature data
                SetParameters(model, feature, featData, parameters);
                feature.ModifyDefinition(featData, model, null);
            }
            else
            {
                //cancel modifications
                featData.ReleaseSelectionAccess();
            }

            return true;
        }

        private bool ShowPage(EditMacroFeatureDefinitionParameters parameters)
        {
            //TODO: Show property page or any other user interface
            return true;
        }
    }
}
