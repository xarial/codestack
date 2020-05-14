---
layout: issue-fix
title: Fix the inconsistent model title extension in SOLIDWORKS API
caption: Model Title Inconsistency Displaying Extension
description: Fixing the Run-time Error '5' - Invalid procedure call or argument error when running a macro which is using the title of the model (e.g. inserting the note, linking the custom property value, generating new file name for exporting)
image: /solidworks-api/troubleshooting/macros/model-title-inconsistency-displaying-extension/invalid-procedure-or-call-error.png
labels: [macro, troubleshooting]
redirect-from:
  - /2018/04/macro-troubleshooting--model-title-inconsistency-displaying-extension.html
---
## Symptoms

SOLIDWORKS macro is using the title of the model (e.g. inserting the note, linking the custom property value, generating new file name for exporting).
As the result macro misbehaves (inserting extension twice) or displays the error: *Run-time Error '5': Invalid procedure call or argument*  

The extension is extracted from the document title via [IModelDoc2::GetTitle](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldoc2~gettitle.html) SOLIDWORKS API method.

{% include img.html src="invalid-procedure-or-call-error.png" width=640 height=211 alt="Run-time Error '5': Invalid procedure call or argument error when running a macro" align="center" %}

## Cause

There are several factors which affect the way title is displayed to the user:

* Extension visibility in the model's title is displayed based on the windows setting *'Hide extension for known file types'*.
Depending on this setting title of the model can either include or exclude extension (e.g. *Part1 *or *Part1.sldprt*)  

{% include img.html src="hide-extensions-for-known-file-types.png" width=277 height=320 alt="Hide extensions for known file types option in Windows explorer" align="center" %}

* For the newly created files (i.e. files which were never saved) extension is never displayed
* For drawings the title is a composition of a name and the active sheet. The extension is never displayed for drawings

## Resolution

* Change the setting based on the macro requirement
* Modify the macro code to consider both options. The example below provides two functions to get the title with or without extension regardless of the conditions.

{% code-snippet { file-name: get-title-extension.vba } %}
