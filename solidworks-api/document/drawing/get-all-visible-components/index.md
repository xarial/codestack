---
layout: article
title: Get all visible components in the drawing view using SOLIDWORKS API
caption: Get All Visible Components
description: VBA macro to get all visible components in the drawing view (including sub-assemblies) using SOLIDWORKS API
image: /solidworks-api/document/drawing/get-all-visible-components/drawing-view-feature-tree.png
labels: [visible components,drawing view]
---
{% include img.html src="drawing-view-feature-tree.png" width=350 alt="Drawing view feature manager tree" align="center" %}

This VBA macro extracts all visible components from the selected drawing view using SOLIDWORKS API. Macro will extract all types of components (part components and assembly components).

[IView::GetVisibleComponents](http://help.solidworks.com/2013/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iview~getvisiblecomponents.html) SOLIDWORKS API methods only extracts part components (i.e. sldprt files) while all sub-assembly components are not returned. Furthermore the pointers to [IComponent2](http://help.solidworks.com/2017/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IComponent2.html) interfaces returned by this function are drawing context components. The [IComponent2::GetParent](http://help.solidworks.com/2016/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.icomponent2~getparent.html) SOLIDWORKS API method returns Nothing for all components which means it is not possible to find the parent sub-assembly.

The below code addresses this limitations and returns all components in the context of their assembly document.

{% include_relative Macro.vba.codesnippet %}
