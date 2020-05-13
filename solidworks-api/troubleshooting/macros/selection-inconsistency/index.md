---
layout: issue-fix
title: Fixing the inconsistent selections in the SOLIDWORKS macro
caption: Selections are inconsistent in the macro
description: Fixing the error when selections in the macro are not consistent
lang: en
image: /solidworks-api/troubleshooting/macros/selection-inconsistency/recorded-macro-extrude.png
labels: [macro, troubleshooting]
---
## Symptoms

SOLIDWORKS macro was recorded by [Macro Recording Tool](http://help.solidworks.com/2012/english/solidworks/sldworks/c_recording_playing_macros.htm) and requires some selections to be made (usually for feature creation or mating). When running the macro the selection may fail or different object is selected which is causing the macro misbehavior.

## Cause

Usually macro recording is using the [IModelDocExtension::SelectByID2](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldocextension~selectbyid2.html) SOLIDWORKS API method to capture the selection. This method may use temporarily names (like sketch segment names) or coordinates for selection which might not be consistent across different models or view orientations.

{% include img.html src="recorded-macro-extrude.png" width=500 alt="Recorded macro line to select arc in the sketch by name" align="center" %}

## Resolution

Update the selection method. Refer the [Selection]({{ "solidworks-api/document/selection" | relative_url }}) article for detailed guide.
