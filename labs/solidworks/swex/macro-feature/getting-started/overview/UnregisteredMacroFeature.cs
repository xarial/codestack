using CodeStack.SwEx.MacroFeature;
using CodeStack.SwEx.MacroFeature.Attributes;
using System.Runtime.InteropServices;

namespace CodeStack.SwEx
{
    [ComVisible(true)]
    [Options("SwExMacroFeature", "CodeStack. Download the add-in")]
    public class UnregisteredMacroFeature : MacroFeatureEx
    {
    }
}
