---
title: Pass Parameters To SOLIDWORKS VBA Macro
caption: Pass Parameters To VBA Macro
description: Workarounds for passing parameters to SOLIDWORKS VBA macro from external applications
labels: [arguments,parameters,interoperability]
---
SOLIDWORKS VBA macros do not accept custom parameters as an input so it is not possible to pass user argument to [ISldWorks::RunMacro2](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~runmacro2.html) method. This limitation could be the major 'roadblock' for performing the automation of SOLIDWORKS using API.

This could be a handy feature in cases where the macro is used as the part of bigger automation where multiple macros need to share the same argument (e.g. output location, time stamp, etc.). Or process is started from the server application or via scheduling software which generates the input which needs to be passed to the macro.  

This section contains the implementation of alternative workarounds to overcome this limitation.

Several options of passing parameters to SOLIDWORKS VBA macros are explored and examples are provided.

* [Via Clipboard](via-clipboard)
* [Via SWB Macro](via-swb-macro)
