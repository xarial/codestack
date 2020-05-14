---
layout: article
title: Render tetrahedron using OpenGL and handle view display modes using SOLIDWORKS API
caption: Render Tetrahedron And Handle View Display Modes
description: VB.NET add-in for SOLIDWORKS to render the graphics of tetrahedron in different display modes (shaded, shaded with edges, hlr, hlv, wireframe)
image: /solidworks-api/adornment/opengl/display-mode-tetrahedron/shaded-with-edges.png
labels: [opengl,display model,shaded,hlr,hlv,wireframe]
---
This example demonstrates how to render the tetrahedron geometry using OpenGL graphics and handle the different display modes (shaded, shaded with edges, hlr, hlv, wireframe).

Once add-in is compiled, tetrahedron will be rendered in each newly opened or created 3D model (part or assembly).

Change display modes in the heads up menu of SOLIDWORKS model view to see the graphics updated.

![Display modes in SOLIDWORKS model view](display-style.png){ width=350 }

## Display Modes

### Shaded with edges

Achieved by rendering two layers of graphics: filled triangles and not filled lines on top of the triangles.

![Shaded with edges display mode](shaded-with-edges.png){ width=200 }

### Shaded

Achieved by rendering the triangles

![Shaded](shaded.png){ width=200 }

### Hidden lines removed

Achieved by rendering triangles with polygon mode set to lines.

![Hidden lines removed display mode](hidden-lines-removed.png){ width=200 }

### Hidden lines visible

Achieved by rendering dashed line in lines mode

![Hidden lines visible display mode](hidden-lines-visible.png){ width=200 }

### Wireframe

Achieved by rendering graphics in lines mode

![Wireframe display mode](wireframe.png){ width=200 }

Source code can be downloaded from [GitHub](https://github.com/codestackdev/solidworks-api-examples/tree/master/swex/add-in/opengl/OglTetrahedron)

## AddIn.vb

This the add-in entry point. [SwEx.AddIn](/labs/solidworks/swex/add-in/) framework is used to manage documents lifecycle by providing the wrapper class.

{% code-snippet { file-name: AddIn.vb } %}

## OpenGlDocumentHandler.vb

This is a handler class for each model document which subscribes to the OpenGL Buffer Swap notification provided by SOLIDWORKS and performs the calculation of tetrahedron triangle coordinates and renders the geometry. 

{% code-snippet { file-name: OpenGlDocumentHandler.vb } %}

## OpenGL.vb

List of imports for OpenGL functions.

{% code-snippet { file-name: OpenGL.vb } %}
