---
caption: Upgrade Legacy
title: Upgrade legacy custom properties to a new architecture
description: VBA macro which upgrade legacy SOLIDWORKS custom properties to a new architecture in SOLIDWORKS 2022
---
This macro upgrades the legacy custom properties to a [new architecture](https://help.solidworks.com/2022/english/solidworks/sldworks/c_custom_properties_architecture.htm) in SOLIDWORKS 2022.

To configure the macro, modify the constant parameters in the macro.

~~~ vb
Const UPDATE_ALL_COMPS As Boolean = True
Const REBUILD_ALL_CONFIGS As Boolean = True
~~~

**UPDATE_ALL_COMPS** sets to rebuild all components of the assembly or top level only
**REBUILD_ALL_CONFIGS** specifies if it is required to rebuild all configurations

{% code-snippet { file-name: Macro.vba } %}
