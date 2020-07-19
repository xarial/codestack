---
layout: sw-tool
title: VBA macro to export sketch point coordinates to CSV file
caption: Export Sketch Coordinates
description: VBA macro to export coordinates of sketch points from the selected sketch with an ability to convert coordinate to user units and into the model space
image: export-coordinates.svg
labels: [sketch,export,points,coordinates,csv]
group: Sketch
---
![Sketch points in the selected sketch](sketch-points.png){ width=500 }

This VBA macro allows to export the coordinates of all sketch points from the selected sketch into the CSV file.

CSV file can be opened in Excel

![Sketch points coordinates opened in Excel](excel-coordinates.png)

Macro has an option to export coordinates in the sketch space (XY for 2D sketch) or in the model space (XYZ). Macro has an option to convert the points coordinates to system units (meters) or user units, currently assigned to the model.

Configure the macro by changing the constants below.

~~~ vb jagged-bottom
Const CONVERT_TO_USER_UNIT As Boolean = True 'True to use the current model units, False to use system units (meters)
Const CONVERT_TO_MODEL_SPACE As Boolean = True 'For 2D Sketches, True to export coordinates in the sketch space, False to convert coordinates to the model space
Const OUT_PATH As String = "D:\points.csv" 'Full path to the output file
~~~

{% code-snippet { file-name: Macro.vba } %}