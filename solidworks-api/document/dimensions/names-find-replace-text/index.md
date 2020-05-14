---
layout: sw-tool
title: Find-replace text in dimension names using SOLIDWORKS API
caption: Find-Replace Text In Dimension Names
description: Macro replaces the text in the dimension names of the selected feature or features
image: /solidworks-api/document/dimensions/names-find-replace-text/rename-dims.png
labels: [dimension, example, find, model, rename, replace, solidworks, solidworks api]
group: Model
redirect-from:
  - /2018/03/find-replace-text-in-dimension-names.html
---
This macro finds and replaces the text in the dimension names of the selected feature or features (similar to Find-Replace feature in text editors) using SOLIDWORKS API:

![Input box for the text to find in the dimension names](rename-dims.png){ width=320 }

1. Open SOLIDWORKS assembly or part
1. Select features to lookup dimensions in
1. Run the macro
1. Specify the text to find and the text to replace. Only include short dimension name.
For example the dimension D1 in Sketch1 will have a short name *D1* and full name *D1@Sketch1.* Specifying *D* in find field and *B* in replace field will result in dimension to be renamed to *B1@Sketch1*.

{% code-snippet { file-name: Macro.vba } %}
