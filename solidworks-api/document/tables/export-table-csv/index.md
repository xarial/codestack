---
layout: sw-tool
title: Export SOLIDWORKS table to CSV using VBA macro
caption: Export Table To CSV
description: Macro exports selected table (BOM, General Table, Revision etc.) into CSV format allowing to export with or without header preserving the special symbols like comma (,) and new line symbol using VBA macro
image: /solidworks-api/document/tables/export-table-csv/export-table-csv.png
logo: /solidworks-api/document/tables/export-table-csv/export-table-csv.svg
labels: [table,csv,export]
categories: sw-tools
group: Model
---
This macro exports the selected table to the CSV (Comma Separated Values) file using SOLIDWORKS API. This functionality is similar to built-in 'Save As' option for table:

{% include img.html src="bom-save-as.png" width=350 alt="Save As option for tables" align="center" %}

However macro preserves the special symbols like commas, quotes or new line symbols and properly escapes them according to the CSV specification:

{% include img.html src="bom-table.png" width=350 alt="Bill Of Materials with special symbols (comma and new line)" align="center" %}

So the file can be later properly read using the CSV readers like MS Excel;

{% include img.html src="bom-table-csv-excel.png" width=350 alt="CSV file imported to Excel" align="center" %}

For the above example BOM table the macro will generate the following output:

~~~ csv
ITEM NO.,PART NUMBER,Description,QTY.
1,B01-A57,Blade shaft,1
2,B01-A12,Top blade,1
3,B02,"Bottom blade
Fixed",1
4,R1284,Blade rivets,4
5,E25-E16,"Blade extension, Plastic",1
~~~

Macro can be configured by modifying the value of the constants

~~~ vb
Const OUT_FILE_PATH As String = "D:\bom.csv" 'Full path to the output CSV file
Const INCLUDE_HEADER As Boolean = True 'True to include the table header, False to only include data
~~~

Specify empty string as the *OUT_FILE_PATH* variable value to export table with the same name as original file into the same folder

For example for the table in the *D:\MyDrawing\Draw001.slddrw* file the below setting would save the file into the *D:\MyDrawing\Draw001.csv* location.

~~~ vb
Const OUT_FILE_PATH As String = ""
~~~


{% code-snippet { file-name: Macro.vba } %}
