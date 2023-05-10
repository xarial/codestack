---
caption: Split Folders To Configurations
title: Split feature folders of the SOLIDWORKS file to individual configurations
description: VBA macro creates individual configurations for each feature folder in the active SOLIDWORKS part or assembly
---
This VBA macro creates configuration for each top-level feature folder in the active SOLIDWORKS part or assembly.

Derived configuration will be created for each feature folder and will be named after it.

All other folders will be suppressed for each configuration. Features outside of the folders will not be suppressed.

{% code-snippet { file-name: Macro.vba } %}