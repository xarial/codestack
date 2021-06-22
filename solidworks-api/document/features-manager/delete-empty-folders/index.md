---
layout: sw-tool
title: Macro to delete all empty feature folders in SOLIDWORKS files
caption: Delete Empty Folders
description: VBA macro deletes all empty feature folders in the SOLIDWORKS files (part or assembly)
image: delete-folders.svg
labels: [feature, empty, delete, cleanup]
group: Model
---
![Delete feature manager folders](delete-folders.svg){ width=300 }

This VBA macro will delete all empty feature folders from the active part or assembly.

> Feature folders which only contain empty folders will also be deleted.

![Empty folders deleted from the feature manager tree](deleted-empty-folders.png)

{% code-snippet { file-name: Macro.vba } %}
