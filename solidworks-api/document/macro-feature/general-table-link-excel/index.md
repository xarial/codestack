---
layout: article
title: SOLIDWORKS macro feature to link and auto update general table to Excel
caption: Link And Auto Update General Table To Excel
description: Macro allows to link and automatically update the general table to external Excel or text/CSV file using SOLIDWORKS API
image: /solidworks-api/document/macro-feature/general-table-link-excel/linked-excel-table.png
labels: [general table, excel, link, macro feature]
---
{% include img.html src="linked-excel-table.png" width=350 alt="Linked table macro feature in the feature tree" align="center" %}

This macro allows to create General Table in part, assembly and drawing and link it to external Excel or text/csv file using SOLIDWORKS API. This macro implemented as embedded macro feature which means that table will be automatically updated once the model is rebuilt.

* Run the macro
* Specify the full path to excel (*.xls or *.xlsx) or comma separated text file (*.csv or *.txt) in the first prompt dialog
* Optionally specify the name of the spreadsheet to read data from. If empty string is specified first spreadsheet will be used

Macro inserts the table and macro feature in the feature tree with the data from external file. Modify the file or general table and rebuild the model - table is updated.

## Notes and limitations

* Only simple CSV files are supported (i.e. simple comma separated values, new line symbols or commas in the values are not supported)
* Excel is not required when CSV file is used
* Using CSV files has significant performance benefits as it is not required to start Excel and load document to get the data. Use this option where applicable
* Excel is displayed invisible and session may be cached for better performance benefits
* If CSV or Excel files are saved relative to the model - relative path will be maintained. It means that the SOLIDWORKS file can be moved together with Excel/CSV and link won't be broken
* If General Table is selected when inserting new feature - this table will be used instead of creating new one
* Currently it is not possible to change the path to external Excel file. Delete the macro feature instead and reinsert it by selecting the general table (see previous point)
* Macro feature is embedded into the model which means that the table will be updated on any other workstations even if this macro is not available.

{% code-snippet { file-name: Macro.vba } %}
