---
layout: article
title: Application frame customization using SOLIDWORKS API
caption: Frame
description: Automating SOLIDWORKS frame (menu, toolbars, command manager) using API
lang: en
image: /images/codestack-snippet.png
labels: [frame,menu,toolbar,commands]
---
Elements displayed in the SOLIDWORKS application frame, such as menu, command manager and tabs, toolbars can be customized using [IFrame](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFrame.html) and [ISldWorks](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks.html) SOLIDWORKS API Interfaces.

In addition frame object provides the access to SOLIDWORKS windows handler via [IFrame::GetHWnd](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iframe~gethwnd.html) method which enables the usage of [Windows API to invoke SOLIDWORKS commands](https://blog.codestack.net/2019/03/solidworks-api-command-doesnt-exist.html).

This section contains examples of using SOLIDWORKS API and Windows API to automate application frame.
