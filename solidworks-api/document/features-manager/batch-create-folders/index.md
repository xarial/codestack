---
caption: Batch Create Folders
title: Batch create feature folders in the active SOLIDWORKS document
description: VBA macro creates specified number of the feature folders with the specified prefix name in the active SOLIDWORKS part or assembly
---
This VBA macro allows to create feature folders in the batch mode in the active SOLIDWORKS assembly or part document.

Macro will ask for the number of folders to be created and the folder prefix name.

Macro will create the specified number of folder with the prefix name followed by the index.

> If folder with the next index already exists, next index will be used for the naming

{% code-snippet { file-name: Macro.vba } %}