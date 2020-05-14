---
layout: sw-tool
title: Export flat pattern view in the drawing using VBA macro
caption: Export Flat Patterns
description: VBA macro to export flat pattern views in the drawing active sheet to DXF or DWG or other format preserving the bend notes, annotations etc. using SOLIDWORKS API
image: /solidworks-api/document/drawing/export-sheet-metal-views/flat-pattern-view.png
labels: [dxf,dwg,export,flat pattern]
categories: sw-tools
group: Drawing
---
{% include img.html src="flat-pattern-dxf.png" height=350 alt="Flat pattern exported to DXF" align="center" %}

This VBA macro exports all flat pattern views from the active sheet in the drawing to the specified format (e.g. DXF or DWG) using SOLIDWORKS API. Macro exports the file to the same folder as original drawing and names files after the drawing view name.

This macro can be used in conjunction with [Rename flat pattern views with cut-list names](/solidworks-api/document/drawing/rename-sheet-metal-views/) macro  if it is required to name exported files after the cut list name.

Specify the output file extension at the beginning of the macro:

~~~ vb
Const OUT_EXT As String = ".dxf"
~~~

## Algorithm

* Traverse all drawing view of the current sheet of the active drawing
* Find all drawing views of flat pattern
* Create new temp drawing and copies the view
* Remove all dimensions
* Remove all tables
* Set view and sheet scale to 1:1
* Fit sheet size to view
* Export to the specified file


{% include_relative Macro.vba.codesnippet %}
