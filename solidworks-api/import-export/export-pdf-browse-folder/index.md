---
caption: Export Drawing As PDF Into Selected Folder
title: Macro to save active drawing as PDF file into selected output folder and close drawing
description: VBA macro which saves active SOLIDWORKS drawing as PDF file to a selected output folder and saves and closes the original drawing
---
This VBA macro performs the following steps with the active SOLIDWORKS drawing:

* Shows *Browse For Folder* dialog to select the output folder for the PDF file
* Saves the active drawing as PDF file into the folder selected in the previous step. File name of the PDF will be the same as file name of the drawing
* If the original drawing was modified, macro saves the changes
* Closes the active SOLIDWORKS drawing document

{% code-snippet { file-name: Macro.vba } %}