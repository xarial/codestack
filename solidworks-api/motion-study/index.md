---
layout: article
title: Automating Motion Study using SOLIDWORKS API
caption: Motion Study
description: Collection of articles and examples related to SOLIDWORKS Motion Study API
image: /solidworks-api/motion-study/motion-study.png
order: 10
---
{% include img.html src="motion-study.svg" width=250 alt="SOLIDWORKS Motion Study API" align="center" %}

SOLIDWORKS Motion Study API provides specific interfaces in the separate [SwMotionStudy](http://help.solidworks.com/2018/english/api/swmotionstudyapi/SolidWorks.Interop.swmotionstudy~SolidWorks.Interop.swmotionstudy_namespace.html) library. It is required to explicitly add the reference to this library to your application if [early binding]({{ "/visual-basic/variables/declaration#early-binding-and-late-binding" | relative_url }}) is needed to be used.

Base interface [IMotionStudyManager](http://help.solidworks.com/2018/english/api/swmotionstudyapi/SolidWorks.Interop.swmotionstudy~SolidWorks.Interop.swmotionstudy.IMotionStudyManager.html) can be accessed via [IModelDocExtension::GetMotionStudyManager](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~GetMotionStudyManager.html) SOLIDWORKS API method.

This section contains macros and code examples for automating Motion Studies in SOLIDWORKS using API.
