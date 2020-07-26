---
layout: sw-macro-fix
title: Fix errors when creating sketch segments using SOLIDWORKS API
caption: Failed to Create Sketch Segments
description: Fixing the inconsistency of sketch segments (line, arcs, etc) or sketch points creation in the macro
labels: [macro, troubleshooting]
redirect-from:
  - /2018/04/macro-troubleshooting-failed-create-sketch-segments.html
---
## Symptoms

SOLIDWORKS macro creates sketch segments (line, arcs, etc) or sketch points using SOLIDWORKS API. And in some cases the elements are not created while it works correct in other cases.

## Cause

By default all entities inserted using the [ISketchManager](https://help.solidworks.com/2016/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchManager.html) interface are created via user interface. That means that the entity cannot be created if the target area (i.e. boundaries of the segments) is not visible in the user interface (e.g. the view is moved or scaled).  

## Resolution

Set [ISketchManager::AddToDB](https://help.solidworks.com/2016/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isketchmanager~addtodb.html) property to *True* before the entities creation and restore the original value once the job is finished.
Setting this option to true will bypass the creation of entities via User Interface rather add the data directly to the model storage. This may also improve performance while creating the entities.
  
{% code-snippet { file-name: add-to-db.vba } %}