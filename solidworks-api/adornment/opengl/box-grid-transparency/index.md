---
title: Render box grid with transparency using OpenGL and SOLIDWORKS API
caption: Render Box Grid With Transparency
description: Rendering grid of cubes with transparency using OpenGL and SOLIDWORKS API
image: opengl-cubes.png
labels: [opengl,render,transparency]
---
![Transparent cubes rendered with OpenGL](opengl-cubes.png){ width=250 }

This example demonstrates how to render cubes in the predefined grid with transparency using OpenGL and SOLIDWORKS API.

Cubes are rendered automatically on all opened 3D models (parts or assemblies).

Parameters can be configured by changing the constants declared in the add-in:

~~~ cs
private const int INST_COUNT = 27;
private const int ROWS_COUNT = 3;
private const int COLUMNS_COUNT = 3;
private const double DIST = 0.05;
private const double WIDTH = 0.1;
private const double HEIGHT = 0.1;
private const double LENGTH = 0.1;
private readonly Color COLOR = Color.FromArgb(200, Color.Blue);
~~~

Note, this approach is a simple method of rendering OpenGL objects, but it doesn't provide the best performance benefits. Refer the [Import XAML File And Render Using VBO](/solidworks-api/adornment/opengl/vbo-xaml-importer/) for a code example of rendering high performance graphics using Open GL Vertex Buffer Object (VBO).

Source code can be downloaded from [GitHub](https://github.com/codestackdev/solidworks-api-examples/tree/master/swex/add-in/opengl/OpenGlBoxGrid)

## AddIn.cs

This the add-in entry point. [SwEx.AddIn](/labs/solidworks/swex/add-in/) framework is used to manage documents lifecycle by providing the wrapper class.

{% code-snippet { file-name: AddIn.cs } %}

## OpenGlDocumentHandler.cs

This is a handler class for each model document which subscribes to the OpenGL Buffer Swap notification provided by SOLIDWORKS and performs the calculation of coordinates of cubes based on the input parameters and renders in the model's graphics view.

{% code-snippet { file-name: OpenGlDocumentHandler.cs } %}

## OpenGL.cs

List of imports for OpenGL functions.

{% code-snippet { file-name: OpenGL.cs } %}
