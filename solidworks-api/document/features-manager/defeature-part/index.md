---
title: Defeature Part (convert to dumb solid) using SOLIDWORKS API
caption: Defeature Part
description: Macro to convert all features in part to dumb solids (defeature part) and surfaces using SOLIDWORKS API
image: part-feature-tree-defeatured.png
labels: [defeature,parasolid]
---
This VBA macro defeatures the active SOLIDWORKS part. Unlike the [Defeature for Part](https://help.solidworks.com/2018/english/solidworks/sldworks/c_defeature_for_parts.htm) functionality, this macro preserves the original geometry and does not simplify it.

Macro copies all solid and surface bodies, deletes all user features and imports the copied bodies using SOLIDWORKS API. Macro will preserve the hidden flag from the original bodies.

**Before:**

![Part with feature tree](part-feature-tree.png){ width=350 }

**After:**
![Part with defeatured tree](part-feature-tree-defeatured.png){ width=350 }

{% code-snippet { file-name: Macro.vba } %}
