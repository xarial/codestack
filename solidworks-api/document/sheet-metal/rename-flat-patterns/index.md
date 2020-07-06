---
layout: sw-tool
title: Rename sheet metal flat patterns features after the cut-list features
caption: Rename Flat Pattern After Cut-Lists
description: VBA macro to rename sheet metal flat patterns after the corresponding cut-list feature names
image: flat-pattern.svg
labels: [cut-list,sheet metal,flat-pattern,rename]
group: Part
---
![Cut-lists and sheet metal flat patterns](renamed-flat-patterns.png){ width=250 }

This VBA macro renames all sheet metal flat pattern features with the name of the corresponding cut-list item.

This macro can be used in conjunction with [Rename Cut List Features](/solidworks-api/document/cut-lists/rename-cut-list-items/) macro.

In order to avoid the name conflict, suffix is added to flat pattern features as below.

~~~ vb jagged-bottom
Const SUFFIX As String = "_FP"
~~~

Macro will automatically add the index to the flat pattern name which shares the same cut list.

Watch [video demonstration](https://youtu.be/jsjN8zNRTuc?t=276)

{% code-snippet { file-name: Macro.vba } %}
