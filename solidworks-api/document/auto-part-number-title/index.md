---
layout: sw-tool
title: Set title as part number for new file using SOLIDWORKS API
caption: Set Title To Next Part Number
description: VBA Macro to set title with automatically incremented number from the shared file using SOLIDWORKS API for new files
image: /solidworks-api/document/auto-part-number-title/automatic-model-title.png
labels: [part number,title]
categories: sw-tools
group: Model
---
{% include img.html src="automatic-model-title.png" width=450 alt="Model title set to part number" align="center" %}

This VBA macro automatically increments the part number and sets this as a title for newly created file using SOLIDWORKS API.

Part number is incremented and stored in the external text file which can be shared across different users if needed.

{% include img.html src="part-number-storage-file.png" width=350 alt="Current part number value in the text file" align="center" %}

Macro provides several options to format the title which can be modified by changes in the values of the constants in the macro.

~~~ vb
Const NMB_SRC_FILE_PATH As String = "D:\prt.txt" 'path to store the current part index
Const NMB_FORMAT As String = "000" 'padding for the number, e.g. 001, 002, instead of 1, 2
Const BASE_NAME As String = "PRT-" 'Base prefix for file naming
~~~

Follow the video tutorial in the [Run Macro On Document Load]({{ "solidworks-api/application/documents/handle-document-load/" | relative_url }}) article for the guide of running this macro automatically for each newly created model.

{% include_relative Macro.vba.codesnippet %}
