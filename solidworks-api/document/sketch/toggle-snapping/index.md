---
layout: sw-tool
caption: Toggle Sketch Snapping
title: Macro to toggle the sketch snapping in SOLIDWORKS document
description: VBA macro to toggle on and off the sketch enable snapping option in SOLIDWORKS sketch
image: toggle-snapping.svg
group: Sketch
---
![Enable Sketch snapping option](enable-snapping-option.png)

This VBA macro allows to toggle on and off the 'Enable' option in SOLIDWORKS sketch.

## Using macro in Toolbar+

This macro can be used in [Toolbar+](https://cadplus.xarial.com/toolbar/) which will improve the user experience. It is possible to enable the [toggle state](https://cadplus.xarial.com/toolbar/configuration/toggles/) for the macro button.

![Enable snapping toggle button](enable-snapping-animation.gif)

Paste this code into the "Toggle Button State Code" text box:

~~~ vb
Return CType(Application, Object).Sw.GetUserPreferenceToggle(249)
~~~

![Code for handling the state of the toggle button](toggle-state-code.png)

Download icon [here](toggle-snapping.svg)

{% code-snippet { file-name: Macro.vba } %}