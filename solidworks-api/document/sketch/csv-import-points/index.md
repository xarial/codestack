---
layout: sw-tool
title: Import points cloud from CSV file into sketch via SOLIDWORKS API
caption: Import Points Cloud From CSV File Into Sketch
description: Macro imports the points cloud from the specified CSV file into the active 2D or 3D Sketch using SOLIDWORKS API
image: /solidworks-api/document/sketch/csv-import-points/points_cloud.png
labels: [csv, points cloud, sketch, import]
categories: sw-tools
group: Sketch
---
{% include img.html src="points_cloud.png" alt="Points cloud in the sketch" align="center" %}

This macro imports the points read from the specified CSV (comma separated values) file into the active sketch using SOLIDWORKS API. Both 2D and 3D Sketches are supported.

First row of the CSV file is considered as a header and ignored. If CSV file doesn't contain the header set the second parameter of *ReadCsvFile* call to *False* in the macro to include first row as a data row.

~~~ vb
vPoints = ReadCsvFile(inputFile, False)
~~~

Coordinate of points are specified in meters.

* [Sample 2D Points Cloud CSV File](points_2d.csv)
* [Sample 3D Points Cloud CSV File](points_3d.csv)

* Open the model and create 2D or 3D sketch
* Run the macro. Specify the full path to CSV file in the displayed prompt window
* Click OK. Points are created in the active sketch

{% include_relative Macro.vba.codesnippet %}
