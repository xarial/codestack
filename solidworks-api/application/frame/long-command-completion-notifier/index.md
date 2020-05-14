---
layout: sw-tool
title: Notify the completion of long running SOLIDWORKS command using SOLIDWORKS API
caption: Notify Long Running Command Completion
description: VBA macro to handle the long running commands in SOLIDWORKS (open, rebuild, suppress etc.) and beep to notify its completion
image: /solidworks-api/application/frame/long-command-completion-notifier/opening-file-progressbar.png
labels: [events,performance,notification,commands]
categories: sw-tools
group: Performance
---
{% include img.html src="opening-file-progressbar.png" width=450 alt="Opening large assembly document in SOLIDWORKS" align="center" %}

This VBA macro will listen for SOLIDWORKS commands (e.g. opening, rebuilding, suppressing, resolving etc.) using SOLIDWORKS API and identify the long running ones by matching the execution time to the user assigned delay period. If the command is running longer then this period, the beep signal is played notifying the user that the command is completed. If command is executed faster than no sound is played.

This can be useful when working with large models as it is possible to switch the screens or perform another activity while command is being executed and be notified once the operation is completed without the need to constantly monitor the progress.

## Running instructions

* Create new macro and add the following code

{% code-snippet { file-name: Macro.vba } %}

* Specify the command minimum delay in seconds by changing the value of *MIN_DELAY* constant
* Create new class module and name it *CommandsListener*. Paste the following code into the class module:
* Start the macro. To automatically start the macro with every SOLIDWORKS session follow the [Run SOLIDWORKS macro automatically on application start]({{ "/solidworks-api/getting-started/macros/run-macro-on-solidworks-start/" | relative_url }}) article.

{% code-snippet { file-name: CommandsListener.vba } %}
