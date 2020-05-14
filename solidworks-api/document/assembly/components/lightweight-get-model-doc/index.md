---
layout: article
title: Get Model Doc from lightweight component using SOLIDWORKS API
caption: Get Model Doc From Lightweight Component
description: Example demonstrates how to get the pointer to IModelDoc2 from the component (even if it is in the suppressed or lightweight state)
image: /solidworks-api/document/assembly/components/lightweight-get-model-doc/lightweight-component.png
labels: [assembly, component, example, lightweight, modeldoc, memory, solidworks api]
---
{% include img.html src="lightweight-component.png" alt="Lightweight component in the assembly tree" align="center" %}

[IComponent2::GetModelDoc2](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IComponent2~GetModelDoc2.html) SOLIDWORKS API method returns the pointer to [IModelDoc2](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2.html) interface.

It is required to use this interface to retrieve the model specific information (such as custom properties, feature tree, annotations etc.).

The model document is not available for the components loaded lightweight or suppressed (i.e. the pointer is *NULL*).

The following example demonstrates how to get the pointer to IModelDoc2 from the component (even if it is in the suppressed or lightweight state) using SOLIDWORKS API. The result is achieved by loading the component directly into memory without the need of resolving the component or opening the file in its own window.

{% include_relative Macro.vba.codesnippet %}
