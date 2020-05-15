---
layout: article
title: Logging capabilities in SwEx framework for SOLIDWORKS add-ins
caption: Logging
description: Logging debug messages from SwEx modules
labels: [logging]
toc-group-name: labs-solidworks-swex
---
All base SwEx modules inherit the [IModule](https://docs.codestack.net/swex/common/html/T_CodeStack_SwEx_Common_Base_IModule.htm) interface which provides an access to [ILogger](https://docs.codestack.net/swex/common/html/T_CodeStack_SwEx_Common_Diagnostics_ILogger.htm) instance allowing to log custom messages and exception from the module.

The following modules provide logger:

* [SwAddInEx](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_SwAddInEx.htm)
* [MacroFeatureEx](https://docs.codestack.net/swex/macro-feature/html/T_CodeStack_SwEx_MacroFeature_MacroFeatureEx.htm)
* [PropertyManagerPageEx](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_PropertyManagerPageEx_2.htm)

Additional options can be specified by decorating the module class via [LoggerOptionsAttribute](https://docs.codestack.net/swex/common/html/M_CodeStack_SwEx_Common_Attributes_LoggerOptionsAttribute__ctor.htm)

{% code-snippet { file-name: LogAddIn.* } %}

Specified logger name will be appended to the SwEx module name (e.g. SwEx.AddIn.MyAddInLog or SwEx.MacroFeature.MyAddInLog or SwEx.PMPage.MyAddInLog).

Log messages are output into the output as setup via [LoggerOptionsAttribute](https://docs.codestack.net/swex/common/html/M_CodeStack_SwEx_Common_Attributes_LoggerOptionsAttribute__ctor.htm) attribute. Currently only debug trace logger is supported. Refer [Troubleshooting](/labs/solidworks/swex/troubleshooting/) article for the instructions of how to capture debug trace messages.