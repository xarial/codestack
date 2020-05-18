---
layout: article
title: Outputting SOLIDWORKS files to different formats using SOLIDWORKS eDrawings API
caption: Output
description: Explanation of options of exporting and printing of SOLIDWORKS file via eDrawings API
image: edrawings-export-types.png
labels: [print,export,edrawings]
---
eDrawings enables exporting of the SOLIDWORKS files to the foreign format listed below:

![Export types in eDrawings](edrawings-export-types.png){ width=450 }

Export can be performed via [IEModelViewControl::Save](https://help.solidworks.com/2016/English/api/emodelapi/eDrawings.Interop.EModelViewControl~eDrawings.Interop.EModelViewControl.IEModelViewControl~Save.html) eDrawings API method.

In addition to the above opened file can be printed via [IEModelViewControl::Print5](https://help.solidworks.com/2016/English/api/emodelapi/eDrawings.Interop.EModelViewControl~eDrawings.Interop.EModelViewControl.IEModelViewControl~Print5.html) eDrawings API method.

Both the exporting and printing APIs are asynchronous. It is required to track the corresponding finish events to find when the process is completed. Use the [OnFinishedPrintingDocument](http://help.solidworks.com/2019/english/api/emodelapi/eDrawings.Interop.EModelViewControl~eDrawings.Interop.EModelViewControl._IEModelViewControlEvents_OnFinishedPrintingDocumentEventHandler.html) and [OnFinishedSavingDocument](http://help.solidworks.com/2019/english/api/emodelapi/eDrawings.Interop.EModelViewControl~eDrawings.Interop.EModelViewControl._IEModelViewControlEvents_OnFinishedSavingDocumentEventHandler.html) for tracking the finishing of printing and saving respectively.

The finish events will not be sent in case of an error. In this scenario it is required to handle the [OnFailedSavingDocument](http://help.solidworks.com/2019/english/api/emodelapi/eDrawings.Interop.EModelViewControl~eDrawings.Interop.EModelViewControl._IEModelViewControlEvents_OnFailedSavingDocumentEventHandler.html) and [OnFailedPrintingDocument](http://help.solidworks.com/2019/english/api/emodelapi/eDrawings.Interop.EModelViewControl~eDrawings.Interop.EModelViewControl._IEModelViewControlEvents_OnFailedPrintingDocumentEventHandler.html) events.