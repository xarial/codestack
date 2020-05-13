---
layout: article
title: Managing user preferences of the document using SOLIDWORKS API
caption: Document Options
description: Collection of articles and examples which demonstrate how to control document options (user preferences) using SOLIDWORKS API
lang: en
image: /images/codestack-snippet.png
labels: [document, preferences, options]
---
To manage user preferences (options) of the SOLIDWORKS part, assembly or drawing it is required to use one of the following SOLIDWORKS API:

For reading the options:

* [IModelDocExtension::GetUserPreferenceDouble](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~GetUserPreferenceDouble.html)

* [IModelDocExtension::GetUserPreferenceInteger](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~GetUserPreferenceInteger.html) 

* [IModelDocExtension::GetUserPreferenceString](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~GetUserPreferenceString.html)

* [IModelDocExtension::GetUserPreferenceTextFormat](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~GetUserPreferenceTextFormat.html)

* [IModelDocExtension::GetUserPreferenceToggle](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~GetUserPreferenceToggle.html)

For writing the options:

* [IModelDocExtension::SetUserPreferenceDouble](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~SetUserPreferenceDouble.html)

* [IModelDocExtension::SetUserPreferenceInteger](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~SetUserPreferenceInteger.html) 

* [IModelDocExtension::SetUserPreferenceString](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~SetUserPreferenceString.html)

* [IModelDocExtension::SetUserPreferenceTextFormat](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~SetUserPreferenceTextFormat.html)

* [IModelDocExtension::SetUserPreferenceToggle](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~SetUserPreferenceToggle.html)

This section contains collection of examples and macros for automating document user preferences using SOLIDWORKS API.