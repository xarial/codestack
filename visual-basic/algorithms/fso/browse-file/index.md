---
title: Show file browse for save or open in Visual Basic 6 (VBA)
caption: Browse File For Save Or Open
description: Displaying file browse dialog to select the save file path or open file path in Visual Basic 6 (VBA)
labels: [files,browse,save]
---
Excel VBA macro provides a helper function to browse the name of the file to save **Application.GetSaveAsFilename** or open **Application.GetOpenAsFilename**. These functions however only available in Excel VBA macros and is not available in other environments.

This example demonstrates how create a generic functions to browse for save or open file.

{% code-snippet { file-name: BrowseFile.vba } %}