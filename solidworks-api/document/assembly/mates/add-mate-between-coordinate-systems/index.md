---
layout: article
title: Add mate between coordinate systems using SOLIDWORKS API
caption: Add Mate Between Coordinate Systems
description: Macro adds coincident mate between 2 coordinate systems of 2 selected components
image: /solidworks-api/document/assembly/mates/add-mate-between-coordinate-systems/sw-mate-coincident.png
labels: [assembly, component, coordinate system, example, mate, solidworks api]
redirect_from:
  - /2018/03/solidworks-api-assembly-add-mate-between-coord-sys.html
  - /solidworks-api/document/assembly/add-mate-between-coordinate-systems
---
Adds coincident mate between 2 coordinate systems of 2 selected components using SOLIDWORKS API. The components must contain the coordinate system features named *Coordinate System1*

{% include img.html src="sw-mate-coincident.png" width=640 alt="Coincident mate property manager page" align="center" %}

[IAssemblyDoc::AddMate3](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iassemblydoc~addmate3.html) SOLIDWORKS API is used to insert mate feature.

{% code-snippet { file-name: Macro.vba } %}
