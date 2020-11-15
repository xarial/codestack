---
layout: sw-tool
title: Open associated drawings of active document or selected components
caption: Open Associated Drawings
description: VBA macro to open the drawings associated to the component in its own window regardless of naming (with an option to open the drawing in detailing mode)
image: open-associated-drawing.svg
labels: [drawing, open, detailing]
group: Drawing
---
This VBA macro allows to open the associated drawings of the selected components in the assembly or active document if nothing is selected.

Unlike out-of-the-box functionality this macro does not have a limitation related to the drawing to be named after the component and located in the same folder. This macro will find all drawings in all sub-folders of the current folder (folder of the active document) regardless if those are named after the component or not.

This macro has an option to open the drawing resolved or in the detailing mode. Modify the value oif **OPEN_DRAWING_DETAILING** to change the behavior.

~~~ vb
Const OPEN_DRAWING_DETAILING As Boolean = True 'opens drawings in detailing mode
~~~

{% code-snippet { file-name: Macro.vba } %}
