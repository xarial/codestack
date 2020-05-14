---
layout: sw-tool
title: VBA macro to review SOLIDWORKS sheets and configurations
caption: Configurations And Sheets Reviewer
description: VBA macro which iterates all sheets and configurations of SOLIDWORKS file and activates each one by one
image: /solidworks-api/document/configuration-reviewer/configurations-reviewer.png
logo: /solidworks-api/document/configuration-reviewer/configurations-reviewer.svg
labels: [configuration,sheet,review,iterate]
categories: sw-tools
group: Model
---
{% include img.html src="configurations.png" alt="Configurations in SOLIDWORKS model" align="center" %}

This VBA macro allows to review all configurations in part or assembly and all sheets in the drawing document of SOLIDWORKS.

Macro will activate each sheet or configuration one by one and wait the specified amount of seconds before activating the next configuration.

Specify the time in seconds to wait before activating next configuration by changing the value of *WAIT_TIME* constant

~~~vb
Const WAIT_TIME As Single = 10 ' wait 10 seconds before activating next configuration or sheet
~~~

Main window will not be blocked so it is possible to manipulate the model in the graphics view.

{% code-snippet { file-name: Macro.vba } %}
