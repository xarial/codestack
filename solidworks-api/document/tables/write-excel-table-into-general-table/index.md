---
layout: sw-tool
title: SOLIDWORKS macro copies data from Excel table into general table
caption: Write Data From Excel Table Into General Table
description: Macro will read all the data from the Excel table and import it into the new general table of the active document or update existing table using SOLIDWORKS API
lang: en
image: /solidworks-api/document/tables/write-excel-table-into-general-table/excel-to-table.png
logo: /solidworks-api/document/tables/write-excel-table-into-general-table/excel-to-table.svg
labels: [table annotation, excel, general table, 2 dimensional array]
categories: sw-tools
group: Model
---
This macro will write the data into the newly created general table of the active document from the specified Excel spreadsheet using SOLIDWORKS API.

Specify the full path to excel file and the name of the spreadsheet in the constants defined in the header of the macro.

In order order to update existing general table instead of creating new one, select the general table in the graphics view or from the feature manager tree and run the macro.

This macro can be embedded into the [Macro Feature]({{ "solidworks-api/document/macro-feature" | relative_url }}) which will allow automatic update of the table. Follow the [Link And Auto Update General Table To Excel]({{ "solidworks-api/document/macro-feature/general-table-link-excel/" | relative_url }}) for more information about this option.

{% include img.html src="excel-table-to-sw-general-table.png" width=500 alt="Excel table with purchase order data imported into SOLIDWORKS General Table" align="center" %}

{% include_relative Macro.vba.codesnippet %}