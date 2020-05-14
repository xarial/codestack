---
layout: article
title: Get bodies and materials from drawing view using SOLIDWORKS API
caption: Get Bodies And Materials From Drawing View
description: VBA macro to find bodies and their materials of the selected drawing view (including sheet metal flat pattern) using SOLIDWORKS API
image: /solidworks-api/document/drawing/get-view-bodies/sheet-metal-views.png
labels: [view bodies,flat pattern]
---
{% include img.html src="sheet-metal-views.png" width=200 alt="Flat pattern drawing views" align="center" %}

This VBA macro finds all bodies of the selected drawing view (including sheet metal flat pattern) and extracts their materials using SOLIDWORKS API.

[IView::Bodies](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iview~bodies.html) property finds the bodies of the drawing view, however this SOLIDWORKS API property returns Nothing for the drawing view created from sheet metal flat pattern.

{% include img.html src="flat-pattern-view-settings.png" height=250 alt="Flat pattern is set in the drawing view property page" align="center" %}

Macro below extracts bodies and finds the materials assigned to them in both cases (for regular parts and for sheet metal patterns). The result is output to Immediate window of VBA editor.

{% include_relative Macro.vba.codesnippet %}
