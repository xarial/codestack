---
layout: article
title: Invoke function of SOLIDWORKS add-in from stand-alone application or macro
caption: Invoke Function Of Add-in
description: Calling function of SOLIDWORKS add-in from stand-alone application or macro (enabling add-in custom API)
labels: [add-in api,invoke]
---
This section contains examples and explains how to create an API for SOLIDWORKS add-in so its functions can be called from [Macros](/solidworks-api/getting-started/macros/), [Stand-Alone Applications](/solidworks-api/getting-started/stand-alone/), [Scripts](/solidworks-api/getting-started/scripts/) or other [Add-Ins](/solidworks-api/getting-started/add-ins/)

Enabling API functions in your add-in might be required when add-in itself needs to be automated. This approach can also help to improve performance. As add-ins are in-process applications, they provide the best performance. In this case add-in can act as an engine for the functionality which gets triggered from the macro or another add-in so the performance is optimal.

There are several approaches could be used to achieve this functionality. Explore the following options for more information:

* [Via Add-In Object](via-add-in-object)