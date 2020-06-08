---
title: Capture measurement of SOLIDWORKS entities using eDrawings markup API
caption: Capturing Measurements
description: Performing measurements capturing of the entities in the SOLIDWORKS model into a text box using eDrawings markup API
image: surveying-form.png
labels: [edrawings,markup,mesurement]
---
![Measurement captured in the text box](surveying-form.png){ width=450 }

This example demonstrate how to use eDrawings markup API to capture the measurements of selected entities into a text box.

This example is based on [Hosting eDrawings control in Windows Forms](/edrawings-api/gettings-started/winforms/)

* Run the form
* Open any SOLIDWORKS or eDrawings file by specifying the full path to a file and clicking *Open* button
* Measurement is automatically enabled
* Select any entity or entities and click *Capture Measurement*. The measurement value is appended into a text box

{% code-snippet { file-name: MainForm.cs } %}

Source code is available on [GitHub](https://github.com/codestackdev/solidworks-api-examples/tree/master/edrawings-api/MeasurementSurveying).
