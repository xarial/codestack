---
caption: Export
title: Export SOLIDWORKS files to foreign formats via Macro+ framework
description: VBA macro demonstrates how to use Macro+ and CAD+ API to export SOLIDWORKS files to multiple specified formats 
image:
macro-plus: vba
---

This VBA macro is [Macro+](https://cadplus.xarial.com/macro-plus/) enabled macro that allows exporting file as to foreign format (e.g. PDF, DWG, STEP, IGES, Parasolid etc.).

Each argument of the macro should specify the output file path and the extension of the exported file.

If specified path is relative than the file will be exported relatively to the input file.

Macro will automatically created directories if needed.

{% code-snippet { file-name: Macro.vba } %}