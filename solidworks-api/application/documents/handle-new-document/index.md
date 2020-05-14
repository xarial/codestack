---
layout: sw-tool
title: Run macro on new document creation using SOLIDWORKS API
caption: Run Macro On New Document Creation
description: VBA macro to run other macros or code on every new document creation using SOLIDWORKS API
image: /solidworks-api/application/documents/handle-new-document/new-document.png
labels: [new document]
categories: sw-tools
group: Model
---
{% include img.html src="new-document.png" width=350 alt="Creating new document in SOLIDWORKS" align="center" %}

This VBA macro handles the creation of new document in SOLIDWORKS (part, assembly or drawing) using SOLIDWORKS API and allows to automatically run custom code or another macro when this event happens. This macro will also handle creation of new virtual document in SOLIDWORKS assembly.

## Configuration

* Create new macro (e.g. *RunOnNewDocCreated.swp*)
* Copy the code into corresponding modules of the macro. The VBA macro tree should look similar to the image below:

{% include img.html src="macro-tree.png" width=250 alt="Macro files tree" align="center" %}

* Place your code into the *main* sub of the *HandlerModule* module. The pointer to [IModelDoc2](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2.html) document is passed as the parameter. Use this pointer instead of [ISldWorks::ActiveDoc](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~activedoc.html) as new document might not be set to active when this event arrives.

~~~ vb
Sub main(model As SldWorks.ModelDoc2)
    'TODO: add your routine here
End Sub
~~~

* It might be useful to automatically run this macro with each session of SOLIDWORKS. Follow the [Run SOLIDWORKS macro automatically on application start]({{ "solidworks-api/getting-started/macros/run-macro-on-solidworks-start/" | relative_url }}) link for more information.
* To learn how to run another macro or group of macros refer [Run Group Of Macros](/solidworks-api/application/frame/run-macros-group/) article

## Macro Module

Entry point which starts new document creation events monitoring

{% code-snippet { file-name: Macro.vba } %}

## FileNewWatcher Class module

Class which handles SOLIDWORKS new document API notifications

{% code-snippet { file-name: FileNewWatcher.vba } %}

## HandlerModule module

Custom VBA code which needs to be run for each newly created document

{% code-snippet { file-name: HandlerModule.vba } %}
