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

Color will be assigned on the feature appearance level.

{% code-snippet { file-name: Macro.vba } %}

## Line Colors

This is an alternative version of the macro which assigns the color as a line color instead of the feature appearance.

This macro will assign the random color for all selected sketches or all sketches if no sketches are selected. **UNABSORBED_ONLY** option is only considered when no sketches are selected.

{% code-snippet { file-name: SetLineColorMacro.vba } %}