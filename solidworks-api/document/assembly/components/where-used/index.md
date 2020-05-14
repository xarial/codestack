---
layout: sw-tool
title: Find the where used assemblies of the selected component using SOLIDWORKS API
caption: Find Where Used
description: VBA macro to find the assemblies within active assembly which are using the selected component using SOLIDWORKS API
image: /solidworks-api/document/assembly/components/where-used/where-used-form.png
labels: [where used,parent,component]
categories: sw-tools
group: Assembly
---
This VBA macro finds all parent components of the selected component instances (Where Used) in the active assembly using SOLIDWORKS API and displays the list for review.

{% include img.html src="where-used-form.png" width=350 alt="Where used form with the list of parent components" align="center" %}

All references can be selected in the form and corresponding component will be highlighted in the Feature Manager Tree.

## Configuration

Macro can be configured by changing the constant parameters at the beginning of the macro as shown below:

~~~ vb
Const CONSIDER_CONFIG As Boolean = False 'True to only find the component which have the same referenced configuration, False to find by model path only
Const INCLUDE_SUPPRESSED As Boolean = False 'True to include suppressed components in the search, False to not include
~~~

## Creating Macro

* Create new macro
* Add new [User Form](/visual-basic/user-forms/)
* Set the name of the form as *WhereUsedForm*
* Drag-n-drop ListBox control onto the form
* Name the ListBox control as *ReferencesList*

{% include img.html src="where-used-form-designer.png" width=550 alt="Form designer" align="center" %}

Place the code into corresponding modules

### Macro

{% code-snippet { file-name: Macro.vba } %}

### WhereUsedForm

{% code-snippet { file-name: WhereUsedForm.vba } %}
