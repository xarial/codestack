---
caption: Export Relative To Coordinate System
title: VBA macro to export active documetn relative to the specified coordinate system using SOLIDWORKS API
description: VBA macro exports active model into the foreign format relative to the specified coordinate system usign SOLIDWORKS API
---

This VBA macro demonstrates how to export active SOLIDWORKS document (part or assembly) to foreign format relative to the specified coordinate system.

File is saved into the same folder with the same name as the original file.

Specify the name of hte coordinate system in **OUT_COORD_SYSTEM_NAME** constant. Specify the output file extension in the **OUT_EXTENSION** constant.

{% code-snippet { file-name: Macro.vba } %}