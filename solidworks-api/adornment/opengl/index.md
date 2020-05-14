---
layout: article
title: Rendering custom 3D graphics with Open GL in SOLIDWORKS API
caption: OpenGL
description: Section contains examples and code snippets for rendering 3D graphics in the model view using SOLIDWORKS API with OpenGL methods
order: 3
image: /images/codestack-snippet.png
labels: [graphics, opengl, render]
redirect-from:

  - /solidworks-api/adornment/open-gl/
---
SOLIDWORKS model is utilizing 3D rendering engines (such as DirectX and OpenGL) to draw the geometry in the graphics area.

SOLIDWORKS API enables 3rd party add-ins and macros to inject 3D graphical objects directly into the model layer using [OpenGL](https://en.wikipedia.org/wiki/OpenGL) functions.

In this case geometry is embedded into the graphics view and transformed together with the model itself. It means that is is not required to calculate the transformation matrix.

OpenGL is the most popular cross-platform library for high performance rendering which is used in gaming, virtual reality, CAD, etc.. Explore the [Documentation](https://www.opengl.org/documentation/) for more information about this library.

OpenGL graphics should be rendered within the [BufferSwapNotify](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dmodelviewevents_bufferswapnotifyeventhandler.html) event of [IModelView](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelView.html) interface.

Explore this section for code examples and macros which demonstrate the usage of OpenGL to render objects in views using SOLIDWORKS API.