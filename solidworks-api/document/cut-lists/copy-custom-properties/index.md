---
layout: sw-tool
title: Macro to copy SOLIDWORKS custom properties from cut-list to model
caption: Copy Cut-List Custom Properties To Model
description: VBA macro to copy SOLIDWORKS custom properties from cut-list (sheet metal or weldment) to model or configuration
image: copy-cutlist-properties.svg
labels: [cut-list,sheet metal,properties]
group: Custom Properties
---
This VBA macro copies the specified or all SOLIDWORKS custom properties from the sheet metal or weldment cut-list item to model or configuration.

Properties from the first found cut-list will be copied.

{% code-snippet { file-name: Macro.vba } %}

## Configuration

Macro can be configured by changing the constants

### Properties Scope

*CONF_SPEC_PRP* constant sets the target properties scope.

* True to copy properties to configuration specific tab
* False to copy to Custom tab

### Properties Source

*COPY_RES_VAL* constant sets the property source

* True to copy resolved values
    
![Resolved values copied to custom properties](copied-property-values.png) { width=500 }

* False to copy expressions

![Expression are copied to custom properties](copied-expressions.png) { width=500 }

### Properties List

*PROPERTIES* array contains list of properties to copy
    
Copy specified properties

~~~ vb
Sub Init(Optional dummy As Variant = Empty)
    PROPERTIES = Array("Prp1", "Prp2", "Prp3") 'Copy Prp1, Prp2, Prp3
End Sub
~~~

Copy all properties

~~~ vb
Sub Init(Optional dummy As Variant = Empty)
    PROPERTIES = Empty
End Sub
~~~