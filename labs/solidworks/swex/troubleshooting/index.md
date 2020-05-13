---
layout: article
title: Troubleshoot SOLIDWORKS add-in developed with SwEx framework
caption: Troubleshooting
description: Troubleshooting techniques for the applications built on SwEx framework
lang: en
image: /labs/solidworks/swex/troubleshooting/debug-view-output.png
toc_group_name: labs-solidworks-swex
order: 7
---
SwEx framework outputs the trace messages which simpifies the troubleshooting process. The messages are output to the default trace listener.

If add-in is debugged from Visual studio than the messages are output to Visual studio Output tab as shown below:

{% include img.html src="visual-studio-output.png" width=450 alt="Trace messages in the output window of Visual Studio" align="center" %}

Otherwise messages can be traced via [DebugView](https://docs.microsoft.com/en-us/sysinternals/downloads/debugview) utility by Microsoft

* Download the utility from the link above
* Unzip the package and run *Dbgview.exe*
* Set the settings as marked below:

Enable *Capture Win32* and *Capture Events* options from the toolbar (marked in red) 
    
{% include img.html src="debug-view-settings.png" width=450 alt="Trace settings in the DebugView utility toolbar" align="center" %}

Alternatively set the capture options via menu as shown below:

{% include img.html src="debug-view-settings-menu.png" width=350 alt="Trace settings in the DebugView utility menu" align="center" %}

Set the filter to filter SwEx messages by clicking the filter button (marked in green)

{% include img.html src="debug-view-filter.png" width=350 alt="Trace settings filter in the DebugView utility" align="center" %}

Messages will be output to trace window

{% include img.html src="debug-view-output.png" width=450 alt="Trace messages in the debug view" align="center" %}

Use *eraser* button to clean messages (marked in blue)

### Notes
* Trace output is very powerful tool for troubleshooting the add-in on clients computers
* DebugView tool is lightweight and doesn't require installation and is provided by Microsoft
* Trace messages will be also output in the release mode
* SwEx framework will output the exception details if thrown while loading of the add-in which can help solving the problem when add-in cannot be loaded

Custom messages and exceptions can be logged from SwEx module. Follow [this link](logging) for more information.