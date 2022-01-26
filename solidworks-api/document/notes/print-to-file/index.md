---
layout: sw-tool
title: Macro to print all notes to the text file from SOLIDWORKS drawing
caption: Print Notes Text To File
description: VBA macro to output all notes text to the text file from the SOLIDWORKS drawing file
image: print-notes.svg
labels: [note, print, regular expression, regex]
group: Drawing
---
This VBA macro outputs text from all drawing views in the SOLIDWORKS drawing to the text file.

Each note will be printed in the new line. It is possible to additionally include the name of the view and the file into the output.

## Configuration

Macro can be configured by modifying the constants

~~~ vb
Const FILE_PATH As String = "" 'Full path to a text file where notes should be written. If empty file is saved with the same name as the original file with _Note.txt prefix
Const PRINT_FILE_NAME As Boolean = True 'True to output the file name to the text file
Const PRINT_VIEW_NAME As Boolean = True 'True to output the drawing view name to the text file
Const FILTER As String = "" 'Regular expression filter for the notes to include (e.g. \d+ to include all notes containing numeric value)
~~~

## Notes

* For the notes which are empty the value will be output as **\[X\]**
* See [Regular Expressions](https://docs.microsoft.com/en-us/dotnet/standard/base-types/the-regular-expression-object-model) for more information about regular expressions which can be used to configure the **FILTER**
* Notes will be appended to an existing text file (new file will be created if not exists). This allows to batch run this macro via [Batch+](https://cadplus.xarial.com/batch/) to output notes from multiple files.

{% code-snippet { file-name: Macro.vba } %}
