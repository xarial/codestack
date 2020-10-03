---
title: Macro to create spline CSV file using SOLIDWORKS API
caption: Create Spline From CSV
description: VBA macro to create spline in the active sketch from the points loaded from the CSV file using SOLIDWORKS API
image: spline-pmpage.png
labels: [csv, sketch, spline]
---
![Spline in the sketch with Property Manager Page](spline-pmpage.png)

This VBA macro demonstrates how to create spline in the active sketch by loading points data from the CSV file. CSV file should contain 3 columns for the coordinates of spline nodes in meters. [Download sample spline data](spline-data.csv)

Specify full path to this file in the **CSV_FILE_PATH** constant

{% code-snippet { file-name: Macro.vba } %}
