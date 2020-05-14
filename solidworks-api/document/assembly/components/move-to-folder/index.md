---
layout: sw-tool
title: Move selected components to feature folder using SOLIDWORKS API
caption: Move To Folder
description: Macro move the components selected in the graphics area into a new folder in the feature manager tree
image: /solidworks-api/document/assembly/components/move-to-folder/move-components-to-folder.png
labels: [components, move to folder]
categories: sw-tools
group: Assembly
---
{% include img.html src="new-folder.png" width=250 alt="Components added to new folder" align="center" %}

This macro allows moving the selected components into the new folder in the feature manager tree using SOLIDWORKS API.

Components (or any of their entities) can be selected in the graphics area. For example only face or edge of the component(s) can be selected for macro to work.

{% include_relative Macro.vba.codesnippet %}