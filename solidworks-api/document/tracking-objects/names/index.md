---
layout: article
title: Reading and changing names of SOLIDWORKS objects (features, components, views) using API
caption: Object Names
description: This article explains the use of object names and the ways to read and change the names
lang: en
image: /solidworks-api/document/tracking-objects/names/face-name.png
labels: [id, track, name]
---
{% include img.html src="face-name.png" width=300 alt="Named face" align="center" %}

Some SOLIDWORKS objects in models can have user names assigned to them. The names are unique identification of the object in the model and it is persistent across rebuild operations or sessions. Names available for viewing and editing from the GUI.

The following object types have names assigned to them

* Component
* Configuration
* Feature
* Layer
* Body
* Sheet
* Dimensions
* Entity (Face, Edge, Vertex)
* Sketch Segment (Line, Arc, Spline, Ellipse)
* Drawing View

### Entity Names

By default names of entities (faces, edges, vertices) are not assigned.

Entity name can be changed from the **Entity Property** dialog. Refer [Displaying Entity Properties](http://help.solidworks.com/2017/english/solidworks/sldworks/hidd_ent_property.htm)

{% include img.html src="entity-property.png" alt="Entity Property dialog box for assigning the entity name" align="center" %}

### Notes and Limitations

* Sketch segment names cannot be changed neither from GUI nor from API

* Names displayed in the selection boxes are not the real names of entities. These are just temporarily assigned names for differentiation the selection in the currently opened property manager page. Those names should not be used as the reference.
{% include img.html src="temp-face-name.png" alt="Temporarily name of face used in the property manager page" align="center" %}

* While changing the name of the component it is required to consider several factors. Refer [Renaming Components]({{ "solidworks-api/document/assembly/components/rename/" | relative_url }}) for more information

The following example allows to rename the selected object with the specified name using SOLIDWORKS API.

{% include_relative Macro.vba.codesnippet %}
