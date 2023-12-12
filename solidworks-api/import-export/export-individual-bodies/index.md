---
caption: Export Individual Bodies
title: Export individual bodies and flat-patterns from SOLIDWORKS part file via Macro+ framework
description: VBA macro demonstrates how to use Macro+ and CAD+ API to export individual bodies to foreign format and flat pattern (for sheet metal) from the active SOLIDWORKS part
image:
macro-plus: vba
---

This VBA macro is [Macro+](https://cadplus.xarial.com/macro-plus/) enabled macro that allows exporting all bodies in the active part file as individual files to foreign format (e.g. STEP, IGES, Parasolid etc.).

Sheet metal bodies could be exported to DXF/DWG format as flat pattern via [Flat Pattern Export](https://cadplus.xarial.com/drawing/export-flat-patterns/) tool API of [CAD+ Toolset](https://cadplus.xarial.com/)

This macro supports the custom argument **bodyName** and it will be resolved to the corresponding body name.

{% code-snippet { file-name: Macro.vba } %}

## CustomVariableValueProvider Class Module

{% code-snippet { file-name: CustomVariableValueProvider.vba } %}