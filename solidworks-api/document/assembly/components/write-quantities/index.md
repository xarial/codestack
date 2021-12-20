---
layout: sw-tool
title: Write component quantity in the SOLIDWORKS assembly to custom property
caption: Write Component Quantity To Custom Property
description: VBA macro which writes the total quantities of components in SOLIDWORKS assembly into custom property
image: bom-quantity.svg
labels: [quantity,component]
group: Assembly
---
This VBA macro calculates the total quantity of each component in the SOLIDWORKS assembly and writes it to the custom property.

This macro can be useful in conjunction with [Export Flat Patterns From Part Or Assembly Components](/solidworks-api/document/sheet-metal/export-all-flat-patterns/) and [Export To Multiple Formats](/solidworks-api/import-export/export-multi-formats/) macros.

## Configuration

Macro can be configured by changing the constant parameters at the beginning of the macro as shown below:

~~~ vb
Const PRP_NAME As String = "Qty" 'Name of the custom property to write quantity to
Const MERGE_CONFIGURATIONS As Boolean = False 'True to consider all configurations of the component as a single item
Const INCLUDE_BOM_EXCLUDED As Boolean = False 'True to write quantities based on the Feature Manager Tree instead of BOM
~~~

## Notes

* Macro will consider the user assigned quantity set via custom property (UNIT_OF_MEASURE)
* Macro will consider configuration BOM options for children components (show, promote or hide)
* Macro will write the quantity property to configuration if **MERGE_CONFIGURATIONS** is set to false and to the document property otherwise
* Macro will not clear existing quantity if it is not in the current scope (for example if component is excluded from BOM)
* Macro will fail for the unloaded components (e.g. lightweight)
* Macro will attempt to resolve all lightweight components

{% code-snippet { file-name: Macro.vba } %}
