---
layout: article
caption: Import To Active Part
title: Macro to import foreign file into active part using SOLIDWORKS API
description: VBA macro to import foreign file (parasolid, step, iges, etc.) directly into the active part document using SOLIDWORKS API
image: imported-file.png
---
![File imported to an active part document](imported-file.png)

This VBA macro demonstrates how to import foreign file with bodies (e.g. parasolid, step, iges, etc.) directly into the active part document.

Change the path to the import file in the **INPUT_FILE** constant

This macro only supports foreign files which are imported as part document.

{% code-snippet { file-name: Macro.vba } %}