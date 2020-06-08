---
title: Passing the parameters to SOLIDWORKS Macro using the SWBasic macro
caption: Via SWBasic Macro
description: Workaround of passing the parameters to the SOLIDWORKS macro via replacing the text in the SWBasic macro
labels: [argument, swb]
---
[SWBasic (*.swb) macros](/solidworks-api/getting-started/macros/types#swbasic-macros.swb) is a legacy type of macros still supported in SOLIDWORKS applications.

One of the benefits of this type of macro is that it is saved in the plain text. This enables 3rd party application to create macros on the fly. In particular this technique can be employed to emulate the passing of parameters to the SOLIDWORKS macro.

For example, the following template macro can be created

**template.swb**

{% code-snippet { file-name: Macro.swb, lang: vba } %}

where *{{Argument1}}* is a placeholder to the argument value to be filled by external application or script:

~~~ cs jagged-bottom 
static void Main(string[] args)
{
    var macroPath = args[0];
    var param = args[1];
    
    var templateMacro = File.ReadAllText(macroPath);
    var macro = templateMacro.Replace("{{Argument1}}", param);

    var tempMacroPath = Path.Combine(Path.GetTempPath(), Path.GetFileName(macroPath));
    File.WriteAllText(tempMacroPath, macro);
~~~

The resulting file can be run as a normal [SOLIDWORKS macro](/solidworks-api/application/frame/run-macros-group/)