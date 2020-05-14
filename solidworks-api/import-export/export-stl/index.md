---
layout: sw-tool
title: Export part or selected component to STL using SOLIDWORKS API
caption: Export Part Or Component To STL
description: Macro exports selected assembly component or part to stl format without the need of activating the document. Macro can optionally apply transformation to the exported STL to reorient the output
image: /solidworks-api/import-export/export-stl/component-stl.png
labels: [component, export, stl]
categories: sw-tools
group: Import/Export
redirect-from:

  - /solidworks-api/import-export/export-component-stl/
---
{% include img.html src="component-stl.png" width=250 alt="Selected component exported to STL" align="center" %}

This C# VSTA macro exports active part or selected component in assembly to STL format using SOLIDWORKS API. Macro will also work with the components loaded lightweight.

This macro is not using the default exporter and overcomes the limitation when the model needs to be loaded in its own window, i.e. [IModelDocExtension::SaveAs](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldocextension~saveas.html) SOLIDWORKS API function is not used. Macro will create stl from the tessellation triangles of the model.

Macro can optionally apply the transform to rotate or move the STL file. It is not required to create a coordinate system for this to happen.

For more information about the STL specification follow [this link](https://en.wikipedia.org/wiki/STL_(file_format)).

## Configuring the orientation

In order to configure the orientation of the output file it is required to change the values of 4x4 orientation matrix defined in the *m_Transform* at the beginning of the macro.

Use the [Get Coordinate System Transform](/solidworks-api/geometry/transformation/get-coordinate-system-transform/) macro to retrieve the transformation from any selected coordinate system.

For example to set the 90 degrees rotation around X axis in clockwise direction it is required to change the values of the *m_Transform* array to the ones below:

~~~ cs
private double[] m_Transform = new double[]
{
    1,-0,0,0,
    0,-1,0,1,
    0,0,0,0,
    1,0,0,0
};
~~~

## Running instructions

* Open part

or

* Open assembly (can be opened lightweight)
* Select part component
* Browse the location of the output STL file
* File is exported

{% code-snippet { file-name: Macro.cs } %}
