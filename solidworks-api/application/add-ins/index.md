---
title: Developing C++, C#, VB.NET add-ins for SOLIDWORKS using API
caption: Add-ins
description: Examples and articles explaining how to work with add-ins in SOLIDWORKS
---
Add-ins are in-process extension to SOLIDWORKS which provide the best performance benefits across all application types. Add-ins are COM objects and must implement the [ISwAddin](http://help.solidworks.com/2012/english/api/swpublishedapi/solidworks.interop.swpublished~solidworks.interop.swpublished.iswaddin.html) interface in SOLIDWORKS API.

Add-ins can be developed with any COM-compatible language: C++, C#, VB.NET, VB6, Managed C++.

Add-ins are available under the Tools->Add-Ins dialog in SOLIDWORKS menu and can be optionally enabled or disabled.

Most of SOLIDWORKS partner products and some of the products of SOLIDWORKS Standard, Professional and Premium packages are developed as add-in application rather than built-in applications.

Add-ins can monitor the full lifecycle of SOLIDWORKS applications and documents. Add-ins have an access to all available SOLIDWORKS API, while macros and stand-alone applications have some limitations as some of the APIs would not be available.