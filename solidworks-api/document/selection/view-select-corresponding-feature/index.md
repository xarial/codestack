---
caption: Select Feature In All Drawing Views
title: Select corresponding features in all drawing views
description: VBA macro which selects the corresponding features of the feature in the model in all drawing views
image: selected-feature.png
---
![Feature selected in the drawing view](selected-feature.png){ width=250 }

This VBA macro demonstrates how to find the pointers for the input feature from the model space in each view in the drawing and select it.

* Open the model drawing views are created from (i.e. assembly or part)
* Select any feature
* Run macro. Macro stops an execution
* Activate drawing
* Continue the macro. All corresponding features are selected in each view

{% code-snippet { file-name: Macro.vba } %}