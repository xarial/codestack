---
caption: Set Configuration BOM Options
title: Macro to change the Bill Of Materials options (Part Number source and children visibility) of SOLIDWORKS configuration
description: VBA macro which changes the Bill Of Materials part number source (configuration name, document name, parent name) and children display (show, hide, promote) options for all or active configuration in SOLIDWORKS document
image: configuration-options.png
---
This VBA macro allows to change the options of the configuration regarding the processing in the Bill Of Materials

* Part Number Source
    * Configuration name
    * Document name
    * Parent name
* Children Components Display
    * Show
    * Hide
    * Promote

![Configuration options Property Manager Page](configuration-options.png)

Macro can process active configuration only or all configurations

Configure the macro by changing its constants

~~~ vb
Const ALL_CONFIGS As Boolean = True 'True to process all configurations, False to process active configuration only
Const PART_NUMBER_SRC As Integer = swBOMPartNumberSource_e.swBOMPartNumber_ConfigurationName 'Part number source: swBOMPartNumberSource_e.swBOMPartNumber_ConfigurationName, swBOMPartNumberSource_e.swBOMPartNumber_DocumentName or swBOMPartNumberSource_e.swBOMPartNumber_ParentName
Const CHILD_COMPS_DISP As Integer = swChildComponentInBOMOption_e.swChildComponent_Promote 'Display of components in BOM: swChildComponentInBOMOption_e.swChildComponent_Show, swChildComponentInBOMOption_e.swChildComponent_Hide or swChildComponentInBOMOption_e.swChildComponent_Promote
~~~

{% code-snippet { file-name: Macro.vba } %}