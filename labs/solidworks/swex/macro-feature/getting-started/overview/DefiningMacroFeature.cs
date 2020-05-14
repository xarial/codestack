using CodeStack.SwEx.MacroFeature;
using CodeStack.SwEx.MacroFeature.Attributes;
using CodeStack.SwEx.Properties;
using SolidWorks.Interop.swconst;
using System;
using System.Runtime.InteropServices;

namespace CodeStack.SwEx
{
    public class MySimplaeMacroFeatureParameters
    {
        public string Parameter1 { get; set; }
    }

    [ComVisible(true)]
    [Guid("47827004-8897-49F5-9C65-5B845DC7F5AC")]
    [ProgId("CodeStack.MyMacroFeature")]
    [Options("MyMacroFeature", swMacroFeatureOptions_e.swMacroFeatureAlwaysAtEnd)]
    [FeatureIcon(typeof(Resources), nameof(Resources.macro_feature_icon), "CodeStack\\MyMacroFeature\\Icons")]
    public class MySimplaeMacroFeature : MacroFeatureEx<MySimplaeMacroFeatureParameters>
    {
    }
}
