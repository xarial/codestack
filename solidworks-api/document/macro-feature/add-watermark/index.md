---
layout: sw-tool
title: Macro feature which adds watermark into SOLIDWORKS model
caption: Add Watermark Feature
description: Adds the watermark feature (license) with custom message and name which cannot be deleted or edited
image: /solidworks-api/document/macro-feature/add-watermark/model-watermark.png
labels: [secutiry, macro feature, lock]
group: Security
---
{% youtube { id: v-S2idKXWDY } %}

This article will explain how to embed a watermark feature in any model using SOLIDWORKS API to protect intellectual property or indicate that the model is to be used under the special conditions. This can be useful if license needs to be embedded into the model which cannot be edited by 3rd parties.

The inserted feature cannot be deleted, suppressed, removed or changed by user interface or API.

The feature can be inserted into any SOLIDWORKS model (part, document, drawing or template).

Feature is fully embedded into the model.

The solution consists of 2 parts:

* Authoring macro. This is used to insert the watermark. This macro is not embedded to the model
* Watermark macro. This macro represents a watermark feature and will be directly embedded into the model

## Setting up authoring macro

* Create new macro in SOLIDWORKS (Tools->Macro->New... command from the SOLIDWORKS menu)
* Specify the name to save this file
* Copy the following code into the macro

{% code-snippet { file-name: InsertWatermark.vba } %}

## Setting up watermark macro

* Create another new macro
* Specify the name of this macro as *Watermark.swp* and save it to the same folder as previous authoring macro

> If different name is used it is required to modify this name in the following constant of the authoring macro

~~~ vb
Const WATERMARK_MACRO_NAME = "Watermark.swp"
~~~

* Change the security note in the constant to be displayed as read-only non-modifiable note.

~~~ vb
Const SECURITY_NOTE = "www.codestack.net"
~~~

* Copy the following code into the macro

{% code-snippet { file-name: Watermark.vba } %}

### Modifying the parameters

There are few parameters which can be modified for the macro which are defined as constants at the top of the macro

~~~ vb
Const MESSAGE As String = "Watermark Feature by CodeStack"
Const FEATURE_NAME As String = "www.codestack.net"
~~~

* *MESSAGE* is the custom message displayed when the watermark feature is edited by the user
* *FEATURE_NAME* name of the watermark feature in the Feature manager tree

### Setting the macro password

Assign the password to the watermark macro so it cannot be changed by users.

* Select *Properties* command from the context menu of Watermark macro

![Properties command in the context menu of VBA macro](vba-macro-properties.png){ width=200 }

* Select the *Protection* tab
* Check the *Lock project for viewing* option
* Specify the password in the password boxes

![VBA macro protection tab](vba-macro-protection.png){ width=250 }

* Save the macro and close

## Inserting the watermark

* To insert the watermark just open the model and run the [Authoring Macro](#setting-up-authoring-macro)
* Save the model

## Watermark behaviour

* Watermark feature will always be moved to the bottom of the feature tree
* When definition of the watermark feature is edited 

![Calling the Edit definition of Watermark feature](vba-edit-watermark-feature.png){ width=250 }

the custom message is displayed

![Custom message box displayed while editing of Watermark feature](watermark-messagebox.png){ width=250 }

* The following message is displayed on attempt to delete the feature

![Deletion of Watermark is impossible](watermark-cannot-be-deleted.png){ width=250 }

* Non editable security note is added in the model at the origin

![Not editable security watermark note](watermark-security-note.png){ width=250 }

* Feature name cannot be changed. It is possible to rename it, but the name will be reverted while state update (e.g. select, rebuild, open model, etc.)