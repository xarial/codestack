---
layout: article
title: Edit feature in the context of the assembly using SOLIDWORKS API
caption: Edit Feature In The Context Of The Assembly
description: Example demonstrates how to modify feature definition in the context of the assembly
image: /solidworks-api/document/features-manager/modify-feature-definition-in-context/edit-extrude-feature-in-context.png
labels: [edit, feature, context]
---
{% include img.html src="edit-extrude-feature-in-context.png" width=450 alt="Boss-Extrude feature is editing in the context of the assembly" align="center" %}

This example demonstrates how to modify feature definition in the context of the assembly using SOLIDWORKS API.

The steps performed in the macro are equivalent of the following steps in SOLIDWORKS User Interface

* Select component of the part which contains extrusion
* Select 'Edit Part' menu in the context menu of the component
* Select extrusion feature and click 'Edit' command from the context menu
* Modify the value of the extrusion in the forward direction
* Click green tick
* Exit the edit part mode

When editing feature in the assembly it is important to follow the correct [Assembly Context]({{ "/solidworks-api/document/assembly/context/" | relative_url }}).

* Example below is implemented as VSTA3 macro
* Select component of the part in the assembly
* Specify the name of the extrude feature as the value of the *EXTRUDE_FEAT_NAME* variable
~~~ cs
const string EXTRUDE_FEAT_NAME = "Boss-Extrude1";
~~~
* Run macro. As the result the value of the extrusion is changed to the value of *EXTRUDE_DEPTH* variable (in meters)
~~~ cs
const double EXTRUDE_DEPTH = 0.02;
~~~

{% include_relative Macro.cs.codesnippet %}