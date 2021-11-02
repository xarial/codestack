---
title: Macro to suspend graphics update using SOLIDWORKS API
caption: Suspend Graphics Update
description: Example demonstrates how to suppress graphics update while performing the operations using SOLIDWORKS API
labels: [api, graphics, utility, suspend, performance]
---
This macro demonstrates how to suspend graphics update while performing operations with feature tree and models (including opening of new documents) using SOLIDWORKS API.

Macro copies the bodies from the external part into the newly created derived configuration of the active part document.

Set the source part path (the part to copy bodies from) via *SRC_PART* constant

~~~ vb
Const SRC_PART As String = "C:\Sample.sldprt"
~~~

Try both options to see the difference by changing the *SUPPRESS_UPDATES* constant

~~~ vb
Const SUPPRESS_UPDATES As Boolean = True 'True to suppress updates, False to show the updates (default behavior)
~~~

Macro performs the following steps

* Opens the model with bodies to copy
* Copies all the bodies into the memory
* Closes the model
* Creates new derived configuration in the original model
* Inserts copied bodies
* Suppresses the created features in all configurations except of this one
* Activates the original configuration

If *SUPPRESS_UPDATES* option is set to true all of the operations will be hidden and only active state of the model will be shown on screen (i.e. model opening, feature insertion etc. will be invisible)

{% code-snippet { file-name: Macro.vba } %}
