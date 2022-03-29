---
layout: sw-tool
caption: Auto Date Custom Property
title: Create a dynamic auto updatable date custom property in SOLIDWORKS file
description: VBA macro which creates a date custom property in SOLIDWORKS file in the specified format with an option to automatically update
image: auto-date-custom-property.svg
group: Custom Properties
---
This VBA macro allows to insert custom property **Date** into file-specific custom property. User has an option to specify the format of the date. Refer [Date and time format string](https://docs.microsoft.com/en-us/dotnet/standard/base-types/standard-date-and-time-format-strings) for more information about supported formats.

## CAD+

This macro is compatible with [Toolbar+](https://cadplus.xarial.com/toolbar/) and [Batch+](https://cadplus.xarial.com/batch/) tools so the buttons can be added to toolbar and assigned with shortcut for easier access or run in the batch mode.

In order to enable [macro arguments](https://cadplus.xarial.com/toolbar/configuration/arguments/) set the **ARGS** constant to true and pass the format as an argument

~~~ vb
#Const ARGS = True
~~~

{% code-snippet { file-name: Macro.vba } %}

This macro can also be embedded into the model to automatically update the date on each rebuild.

{% code-snippet { file-name: MacroFeature.vba } %}