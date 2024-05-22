---
caption: Get Note Text Format
title: VBA macro to get formatting text form the selected SOLIDWORKS note
description: VBA macro puts the formatted note text (including the font parameters, size and color) from the selected note in the SOLIDWORKS document into the clipboard
image: note-format-text.png
---
![Formatted note text](note-format-text.png){ width=800 }

This VBA macro copies the value of formatted text from the selected note in SOLIDWORKS part, assembly or drawing and copies the value to the clipboard.

Formatted note text includes font information (size, style, color), align, paragraph properties, etc.

![Note formatting](note-formatting.png){ width=800 }

{% code-snippet { file-name: Macro.vba } %}