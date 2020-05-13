---
layout: article
title: Hosting custom user controls in SOLIDWORKS panels using SwEx.AddIn framework
caption: Hosting User Controls
description: Hosting custom user controls in SOLIDWORKS panels (task pane, model view manager, feature manager, options dialog) using SwEx.AddIn framework
lang: en
toc_group_name: labs-solidworks-swex
order: 4
---
Frameworks simplifies adding and managing of [custom user controls](https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.usercontrol?view=netframework-4.8) in the standard panels of SOLIDWORKS.

* Task Pane - application scope panel (usually located on the right hand side of SOLIDWORKS window)
* Model View Manager - document scope panel (usually located at the model of document window). Used to control custom model view (such as motion views, FEA, etc)
* Feature Manager View - document scope panel (tab of the Feature Manager Design Tree, usually located on the right hand side of SOLIDWORKS document). Used to add custom feature tree elements, such as electrical tree, costing, architectural etc.
* Settings - custom page in the SOLIDWORKS settings dialog