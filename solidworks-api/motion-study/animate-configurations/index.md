---
layout: sw-tool
title: Macro animates switching of configurations using SOLIDWORKS API
caption: Animate Configurations
description: Macro demonstrates how to create an animation from configurations to represents model history or sheet metal folding
lang: en
image: /solidworks-api/motion-study/animate-configurations/motion-study-configuration-animation.png
labels: [motion, animation, sheet metal, bending]
categories: sw-tools
group: Motion Study
---
{% include youtube.html id="t35Kjjq509w" width=560 height=315 %}

Macro demonstrates how to create an animation from configurations using SOLIDWORKS API.

This could be useful when it is required to create an animation to represents model history or sheet metal folding.

* Open part or assembly
* Select configurations in the order they should be animated

{% include img.html src="sheet-metal-bending-animation.png" width=350 alt="Multiple configurations selected in the configurations tab" align="center" %}

* Run the macro. New assembly created with configurations set as animation steps.

{% include img.html src="motion-study-configuration-animation.png" width=450 alt="Sheet metal bending animation" align="center" %}

Macro parameters (time of the bend transition and pause between folding operations) can be changed by modifying the constants at the top of the macro

~~~ vb
Const TRANSITION_TIME As Double = 0.5
Const PAUSE_TIME As Double = 2
~~~

Refer the [Suppress Features In New Configurations]({{ "solidworks-api/document/features-manager/create-feature-configurations/" | relative_url }}) for a macro to create configurations from features.

{% include_relative Macro.vba.codesnippet %}