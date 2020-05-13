---
layout: sw-tool
title: Link Cut-List Custom Properties To File With SOLIDWORKS Macro Feature API
caption: Link Cut-List Custom Properties To File Custom Properties
description: Macro feature to link specified custom properties from weldment cut-lists to SOLIDWORKS file custom properties
lang: en
image: /solidworks-api/document/macro-feature/link-cut-list-properties/cut-list-link-macro-feature.png
labels: [macro feature,cut-list,link properties]
categories: sw-tools
group: Part
---
{% include img.html src="linked-custom-properties.png" width=450 alt="Linked file custom properties" align="center" %}

This VBA macro inserts the macro feature using SOLIDWORKS API into the part file which allows to dynamically link specified cut-list custom properties to the file generic custom properties.

{% include img.html src="cut-list-properties.png" width=250 alt="Cut-List custom properties" align="center" %}

Macro feature rebuilds automatically when the parent weldment feature (e.g. structural member feature) is changed. Regeneration method is handling the post update notification which allows to read the up-to-date values of cut-list custom properties.

> Reading the custom properties directly from the swmRebuild function will not return the up-to-date values as at the moment of the regeneration all the properties are not evaluated yet.

Macro feature is inserted into the feature tree and can be suppressed or removed.

One of the benefits of this approach comparing to linking the properties directly with the equation is that it is not name dependent, i.e. properties will remain linked even if cut-list renamed (for example when structural member profile is changed).

{% include img.html src="cut-list-link-macro-feature.png" width=250 alt="Macro feature in the feature manager tree" align="center" %}

## Instructions

* Create new macro and copy the code below

{% include_relative Macro.vba.codesnippet %}

* Add new class module to the macro and name it *PostRegenerateListener*. Place the code below into the class module

{% include_relative PostRegenerateListener.vba.codesnippet %}

* Configure the properties which needs to be linked in the *Class_Initialize* function in *PostRegenerateListener*

~~~ vb
Private Sub Class_Initialize()
    LinkedProperties = Array("DESCRIPTION", "LENGTH", "QUANTITY", "Another Property", "...")
End Sub
~~~

* Select the weldment feature (e.g. structural member) and run the macro. Macro feature is inserted and embedded into the model. You can close and reopen model and SOLIDWORKS session - feature will automatically rebuild. Model can be shared with other users and the behavior will be preserved.
