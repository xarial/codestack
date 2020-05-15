---
layout: sw-tool
title: Set multiple assembly components solving (rigid or flexible) using SOLIDWORKS API
caption: Set Components Solving (Rigid or Flexible)
description: VBA macro to batch set the rigid or flexible option for selected components in the assembly using SOLIDWORKS API
image: batch-set-solving.png
labels: [batch,solving,rigid,flexible]
group: Assembly
---
![Setting the solving for multiple assembly components](batch-set-solving.png)

User can change the solving options (rigid or flexible) for assembly components using components options page or toolbar command. This is however only limited for one component at a time.

![Solving options for the components page](solving-options.png)

This VBA macro allows to set either rigid or solved options for all selected components as one command using SOLIDWORKS API.

Specify the option as follows:

~~~ vb
Const SET_FLEXIBLE As Boolean = True 'True - set to flexible, False - set to Rigid
~~~

{% code-snippet { file-name: Macro.vba } %}
