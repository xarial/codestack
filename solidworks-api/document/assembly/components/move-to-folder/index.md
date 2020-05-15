---
layout: sw-tool
title: Move selected components to feature folder using SOLIDWORKS API
caption: Move To Folder
description: Macro move the components selected in the graphics area into a new folder in the feature manager tree
image: move-components-to-folder.png
labels: [components, move to folder]
group: Assembly
---
![Components added to new folder](new-folder.png){ width=250 }

This macro allows moving the selected components into the new folder in the feature manager tree using SOLIDWORKS API.

Components (or any of their entities) can be selected in the graphics area. For example only face or edge of the component(s) can be selected for macro to work.

{% code-snippet { file-name: Macro.vba } %}