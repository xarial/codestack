---
layout: sw-tool
title: Macro animates switching of configurations using SOLIDWORKS API
caption: Animate Configurations
description: Macro demonstrates how to create an animation from configurations to represents model history or sheet metal folding
image: animate-configurations.svg
labels: [motion, animation, sheet metal, bending]
group: Motion Study
---
{% youtube { id: t35Kjjq509w } %}

Macro demonstrates how to create an animation from configurations using SOLIDWORKS API.

This could be useful when it is required to create an animation to represents model history or sheet metal folding.

* Open part or assembly
* Select configurations in the order they should be animated

![Multiple configurations selected in the configurations tab](sheet-metal-bending-animation.png){ width=350 }

* Run the macro. New assembly created with configurations set as animation steps.

![Sheet metal bending animation](motion-study-configuration-animation.png){ width=450 }

Macro parameters (time of the bend transition and pause between folding operations) can be changed by modifying the constants at the top of the macro

~~~ vb
Const TRANSITION_TIME As Double = 0.5
Const PAUSE_TIME As Double = 2
~~~

Refer the [Suppress Features In New Configurations](solidworks-api/document/features-manager/create-feature-configurations/) for a macro to create configurations from features.

{% code-snippet { file-name: Macro.vba } %}