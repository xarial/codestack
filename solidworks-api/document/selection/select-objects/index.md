---
layout: article
title: Select any SOLIDWORKS objects in a batch using API
caption: Select Any SOLIDWORKS Objects In A Batch
description: Example demonstrates how to select any SOLIDWORKS objects (entities, features, annotations, etc.) in a batch mode
image: /solidworks-api/document/selection/select-objects/select-objects.png
labels: [selection, batch selection, dispatch]
---
![Different object types selected in the graphics area]( select-objects.png)

This example demonstrates how to select any SOLIDWORKS objects (entities, features, annotations, etc.) in a batch mode.

This technique can be useful when the type of the object is not known in advance. It also gives performance benefits when selecting several objects at a time instead of selecting one-by-one using SOLIDWORKS API.

The following example provides similar functionality to SOLIDWORKS [Create Selection Set](http://help.solidworks.com/2015/english/whatsnew/t_creating_selection_sets.htm)

![Create Selection Set context menu command](create-selection-set.png){ height=300 }

* Open any model and select any objects (this can be different types objects like features, entities, annotations etc.)
* Run the macro. Macro will collect the pointers of all selected object
* Macro clears the selection and stops the execution
* Continue the execution and all previously selected objects are reselected.

<details open>
<summary>VBA Example</summary>
{% code-snippet { file-name: Macro.vba } %}
</details>

<details open>
<summary>C# Example</summary>
{% code-snippet { file-name: vsta-macro.cs } %}
</details>
