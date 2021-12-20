---
layout: sw-tool
title: Macro to colorize SOLIDWORKS sheet metal and weldment cut-list items
caption: Colorize Cut Lists
description: SOLIDWORKS VBA macro to colorize all the cut-list item bodies (sheet metal and weldments) based on the value of the custom property
image: color-cut-list.svg
labels: [cut-list,sheet metal,weldment,color]
group: Cut-List
---
This VBA macro allows to assign a unique color for each group of cut-list items (sheet metal or weldment) based on the value of the custom property.

The most common use of this macro will be to differentiate different type of weldment items based on the profile size.

Macro will automatically assign random color to the specific group. It is possible to specify the constant colors to use for the specific group instead of random colors.

## Configuration

In order to specify the name of the custom property to read the value from and group cut-list items, change the value of the **PRP_NAME** constant

~~~ vb
Const PRP_NAME As String = "Description" 'Change the value of Description to select different custom property
~~~

In order to specify colors it is required to modify the values within the **InitColors** method.

~~~ vb
Sub InitColors(Optional dummy As Variant = Empty)

    ColorsMap.Add "SB BEAM 80 X 6", RGB(255, 0, 0)
    ColorsMap.Add "TUBE, RECTANGULAR 50 X 30 X 2.60", RGB(0, 255, 0)
    
End Sub
~~~

To add new color to the map add the following line

~~~ vb
ColorsMap.Add "[PROPERTY VALUE]", RGB([Red], [Green], [Blue])
~~~

For example to add the blue (RGB = 0, 0, 255) color to the weldment profile "50 X 50", it is required to add the following line

~~~ vb
ColorsMap.Add "50 X 50", RGB(0, 0, 255)
~~~

{% code-snippet { file-name: Macro.vba } %}
