---
layout: sw-tool
title: Macro feature to automatically run SOLIDWORKS macro on rebuild
caption: Automatically Run Macro On Rebuild
description: Macro allows to automatically run macros on every rebuild using the macro feature and designed binder attachment with SOLIDWORKS API
lang: en
image: /solidworks-api/document/macro-feature/run-macro-on-rebuild/design-binder-macro-attachment.png
labels: [auto run, macro, rebuild]
categories: sw-tools
group: Model
redirect_from:
  - /solidworks-api/document/run-macro-on-rebuild/
---
This macro allows to automatically run specified macro with every rebuild operation using SOLIDWORKS API.

To setup the macro:

* Run the macro
* Specify the full path to the macro to run
{% include img.html src="input-macro-file-path.png" width=250 alt="Select path to the macro for running" align="center" %}

This macro will be added to the model as a design binder attachment
{% include img.html src="design-binder-macro-attachment.png" width=250 alt="Macro added as a design binder attachment" align="center" %}
* Macro feature is inserted to the macro as the last feature in the tree.

When model is rebuilt (either on demand or automatically) macro will be automatically run.

By default macro feature and a macro are embedded into the model. That means that the model can be opened and will be updated on any other workstation which doesn't have this macro available.

This can be also embedded directly to the document template.

### Options
Macro feature name can be changed via constant

~~~ vb
Const BASE_NAME As String = "[Name of Feature]"
~~~

By default the macro is embedded into the model. In order to edit the macro code use the *Edit* command in the design binder attachment

{% include img.html src="edit-embeded-macro.png" width=250 alt="Edit embedded macro in the design binder" align="center" %}

In order to avoid embedding of the macro change the following constant to *False*

~~~ vb
Const EMBED_MACRO As Boolean = False
~~~

{% include_relative Macro.vba.codesnippet %}
