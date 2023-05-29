---
caption: Split Folders To Configurations
title: Split feature folders of the SOLIDWORKS file to individual configurations
description: VBA macro creates individual configurations for each feature folder in the active SOLIDWORKS part or assembly
---
This VBA macro creates configuration for each top-level feature folder in the active SOLIDWORKS part or assembly.

If no objects selected in the model then all folder features will be processed, otherwise only selected feature folders will be processed.

Created configuration will be named after the feature folder.

It is possible to specify if derived or top level configurations should be created for each feature folder.

~~~ vb
Const CREATE_DERIVED_CONFS As Boolean = True 'True to create derived configuration, False to create top level configuration
~~~

All other folders will be suppressed for each configuration. Features outside of the folders will not be suppressed.

{% code-snippet { file-name: Macro.vba } %}