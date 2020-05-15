---
layout: sw-tool
title: Defeature Part (convert to dumb solid) using SOLIDWORKS API
caption: Defeature Part
description: Macro to convert all features in part to dumb solids (defeature part) and surfaces using SOLIDWORKS API
image: part-feature-tree-defeatured.png
labels: [defeature,parasolid]
group: Part
---
This macro emulates the functionality of [Defeature for Part](http://help.solidworks.com/2018/english/solidworks/sldworks/c_defeature_for_parts.htm) but not using it directly.

Macro copies all visible solid and surface bodies, deletes all user features and imports the copied bodies using SOLIDWORKS API.

**Before:**

![Part with feature tree](part-feature-tree.png){ width=350 }

**After:**
![Part with defeatured tree](part-feature-tree-defeatured.png){ width=350 }

{% code-snippet { file-name: Macro.vba } %}
