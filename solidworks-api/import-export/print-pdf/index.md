---
caption: Print To PDF
title: Print active SOLIDWORKS drawing to PDF printer using SOLIDWORKS API
description: VBA macro to print active SOLIDWORKS drawing to PDF printer using SOLIDWORKS API
---

This VBA macro prints active SOLIDWORKS drawing into the PDF printer. Set the printer name in **PRINTER_NAME** constant.

> This macro displays Save As dialog if target file cannot be overwritten (e.g. opened in another application or read-only)

{% code-snippet { file-name: Macro.vba } %}