---
layout: article
title: Utilizing markup functionality using SOLIDWORKS eDrawings API
caption: Markup
description: Guide on using the markup functionality (measurements, stamps, comments) using eDrawings API
image: /images/codestack-snippet.png
labels: [edrawings,markup,getting started]
---
eDrawings markup API (such as comments, stamps, measurements) can be accessed via [IEModelMarkupControl](http://help.solidworks.com/2016/english/api/emodelapi/eDrawings.Interop.EModelMarkupControl~eDrawings.Interop.EModelMarkupControl.IEModelMarkupControl.html) interface.

Interop can be found in the eDrawings installation folder: *%commonprogramfiles%\eDrawings[Version]\eDrawings.Interop.EModelMarkupControl.dll*

Markup interface can be accessed by calling the [IEModelViewControl::CoCreateInstance](http://help.solidworks.com/2018/english/api/emodelapi/eDrawings.Interop.EModelViewControl~eDrawings.Interop.EModelViewControl.IEModelViewControl~CoCreateInstance.html) eDrawings API method.

It is possible to pass both version specific and version independent GUID or ProgId of the markup control.

Version independent guid can be located in the registry *HKEY_CLASSES_ROOT\EModelViewMarkup.EModelNonVersionSpecificMarkupControl\CLSID*

{% include img.html src="non-version-specific-markup-guid.png" alt="Version independent eDrawings Markup control GUID" align="center" %}

Version specific guids can be located under the corresponding version of the markup control (e.g. *EModelViewMarkup.EModelViewMarkupControl.18* for *eDrawings 2018* or *EModelViewMarkup.EModelViewMarkupControl.19* for *eDrawings 2019*)

~~~ cs
//creating version independent instance of markup using prog id
var eDrawingsMarkupCtrl = eDrawingsCtrl.CoCreateInstance("EModelViewMarkup.EModelMarkupControl") as EModelMarkupControl;
...
//creating version independent instance of markup using guid
var eDrawingsMarkupCtrl = eDrawingsCtrl.CoCreateInstance("{5BBBC05A-BD4D-4e3b-AD5B-51A79DFC522F}") as EModelMarkupControl;
...
//creating version specific instance of markup (eDrawings 2018) using prog id
var eDrawingsMarkupCtrl = eDrawingsCtrl.CoCreateInstance("EModelViewMarkup.EModelMarkupControl.18") as EModelMarkupControl;
~~~