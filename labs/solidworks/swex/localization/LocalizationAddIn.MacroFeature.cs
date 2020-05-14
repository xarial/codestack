[Title(typeof(Resources), nameof(Resources.MacroFeatureBaseName))]
[ComVisible(true)]
public class LocalizedMacroFeature : MacroFeatureEx
{
    protected override MacroFeatureRebuildResult OnRebuild(ISldWorks app, IModelDoc2 model, IFeature feature)
    {
        if (!string.IsNullOrEmpty(model.GetPathName()))
        {
            return MacroFeatureRebuildResult.FromStatus(true);
        }
        else
        {
            return MacroFeatureRebuildResult.FromStatus(false, Resources.MacroFeatureError);
        }
    }
}