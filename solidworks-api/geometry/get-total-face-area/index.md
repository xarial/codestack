---
caption: Get Total Face Area
title: VBA macro to get total face area of the part file
description: VBA macro to calculate total face area of all faces from all bodies (including surface bodies) and display in the user units
---

This VBA macro finds the total area of all faces of all bodies (optionally only visible bodies) in the active part file. Macro will consider both solid and surface bodies.

Results is displayed in the message box in the user units.

{% code-snippet { file-name: Macro.vba } %}