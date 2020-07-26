---
title: Automating Motion Study using SOLIDWORKS API
caption: Motion Study
description: Collection of articles and examples related to SOLIDWORKS Motion Study API
image: motion-study.svg
order: 10
---
![SOLIDWORKS Motion Study API](motion-study.svg){ width=250 }

SOLIDWORKS Motion Study API provides specific interfaces in the separate [SwMotionStudy](https://help.solidworks.com/2018/english/api/swmotionstudyapi/SolidWorks.Interop.swmotionstudy~SolidWorks.Interop.swmotionstudy_namespace.html) library. It is required to explicitly add the reference to this library to your application if [early binding](/visual-basic/variables/declaration#early-binding-and-late-binding) is needed to be used.

Base interface [IMotionStudyManager](https://help.solidworks.com/2018/english/api/swmotionstudyapi/SolidWorks.Interop.swmotionstudy~SolidWorks.Interop.swmotionstudy.IMotionStudyManager.html) can be accessed via [IModelDocExtension::GetMotionStudyManager](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~GetMotionStudyManager.html) SOLIDWORKS API method.

This section contains macros and code examples for automating Motion Studies in SOLIDWORKS using API.
