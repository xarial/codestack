---
layout: article
title: Generate box geometry (solid, sheet, wire) Macro Feature using SOLIDWORKS API
caption: Generate Box Geometry
description: Example of creating VBA macro feature which generates different types of box geometry (solid, sheet, wire) using SOLIDWORKS API
image: /solidworks-api/document/macro-feature/geometry/solid-body.png
labels: [macro feature,geometry,box,solid,sheet,wire]
---
This VBA example demonstrates how to create macro feature which generates custom geometry.

Open part document and run the macro. New feature is inserted in the Feature Manager tree and box geometry is generated either as solid, sheet or wire body.

## Configuration

### Embedding

Set the value of *EMBED_MACRO_FEATURE* constant to specify if macro feature should be embedded to file or not. If this option set to *True* then part document can be opened on any other computer and the geometry will be present without the need to copy the macro.

### Box Size

Size of the box can be configured by changing the *WIDTH*, *LENGTH* and *HEIGHT* constants:

~~~ vb
Const WIDTH As Double = 0.01
Const LENGTH As Double = 0.01
Const HEIGHT As Double = 0.01
~~~

### Body Type

Generated body type can be set by assigning the value to *BODY_TYPE* constant

#### swBodyType_e.swSolidBody

Creates a box as a solid body geometry.

{% include img.html src="solid-body.png" width=350 alt="Macro feature generates solid body" align="center" %}

#### swBodyType_e.swSheetBody

Creates a single surface body by sewing the faces of the box.

{% include img.html src="surface-body.png" width=350 alt="Macro feature generates surface (sheet) body" align="center" %}

#### swBodyType_e.swWireBody

Creates wire bodies from all edges of the box geometry. Wire bodies are edges and not presented in the bodies folders. Example of wire bodies used in standard feature tree are curves (composite, through XYZ, projected etc.)

{% include img.html src="wire-body.png" width=350 alt="Macro feature generates wire body" align="center" %}

{% code-snippet { file-name: Macro.vba } %}
