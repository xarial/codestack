---
layout: article
title: Create extruded slot temp body using SOLIDWORKS modeler API
caption: Create Extruded Slot Temp Body
description: Example demonstrates how to extrude the slot profile to create a temp body using SOLIDWORKS API and IModeler interface
image: /solidworks-api/geometry/primitives/create-extruded-slot/extruded-slot.png
labels: [topology, geometry, extrude, slot]
---
![Extruded slot profile](extruded-slot.png){ width=250 }

This VBA example demonstrates how to create a temp body by extruding the slot profile.

Macro will stop the execution and display the preview of the slot in the graphics area. Continue the macro to hide the preview and dispose temp body.

Slot profile is built in the *GetSlotProfileBody* function as per the parameters below:

![Parameters of the slot](slot-parameters.png){ width=250 }

{% code-snippet { file-name: Macro.vba } %}
