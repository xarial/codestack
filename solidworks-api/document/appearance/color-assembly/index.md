---
caption: Apply Random Colors To Components
title: Macro to apply random colors to components in SOLIDWORKS assembly
description: VBA macro to apply random color to all components in the SOLIDWORKS assembly with an option to apply on a component or model level and group by custom property value
---
This VBA macro applies a random color on all components of the active assembly.

Modify constants of the macro to change the level of the color (component or model level).

If colors is applied to the individual configurations (e.g. **ALL_CONFIGS** = **False**), documents must have a display state linked to the configuration, otherwise the color cannot be configuration specific

~~~ vb
Const COMP_LEVEL As Boolean = True 'True to apply color on the assembly level, False to apply color on a model level
Const PARTS_ONLY As Boolean = True 'True to only process part components, False to apply color to assemblies as well
Const ALL_CONFIGS As Boolean = True 'True to apply color to all configurations, False to apply to referenced configuration only
~~~

~~~ vb
Const PRP_NAME As String = "Type" 'Custom property to group color by, Empty string "" to not group components

Sub InitColors(Optional dummy As Variant = Empty)

    ColorsMap.Add "Plate", RGB(255, 0, 0) 'Color all component which custom property 'Type' equals to 'Plate' to Red color
    ColorsMap.Add "Beam", RGB(0, 255, 0) 'Color all component which custom property 'Type' equals to 'Beam' to Green color
    
End Sub
~~~

{% code-snippet { file-name: Macro.vba } %}