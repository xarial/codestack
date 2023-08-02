---
caption: Assign Random Color To Sketches
title: Macro to assign random color to sketches in the document
description: VBA macro assigns random color to sketches in SOLIDWORKS parts or assemblies with an option to skip already assigned sketches and unabsorbed sketches
---

This VBA macro assigns the random color to all sketches of active parts or assemblies.

Macro can be configured to skip sketches with already assigned colors and select only unabsorbed sketches (e.g. sketches which are not used in other features)

~~~vb
Const SKIP_ASSIGNED As Boolean = False 'Processes all sketches (including the sketches with assigned colors)
Const UNABSORBED_ONLY As Boolean = False 'Process all sketches (absorbed and unabsorbed)
~~~

{% code-snippet { file-name: Macro.vba } %}