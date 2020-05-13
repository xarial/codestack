---
layout: article
title: SOLIDWORKS API Object Model class hierarchy diagram
caption: Class Diagram Relationship
description: Detailed explanation of relationships (class diagram) between classes and interfaces in the SOLIDWORKS API Objects Model
lang: en
image: /solidworks-api/getting-started/api-object-model/class-diagram/class-diagram.png
labels: [hierarchy, classes, model]
order: 1
javascripts:
    - "scripts/svg-pan-zoom.min.js"
---
Diagram below displays the relationship between interfaces in the SOLIDWORKS API Object model. This is not a complete class hierarchy rather the representation of the mostly commonly used methods and interfaces.

Diagram is interactive, it is possible to zoom in and zoom out with mouse wheel as well as pan with right or left mouse button. Navigation control in the bottom right corner allows to zoom in and zoom out as well as zoom to fit

{% include img.html src="control-box.png" alt="Control Box" align="center" %}

All boxes and arrows are clickable and will redirect for the information page about particular method, property or interface.

{% include img.html src="open-link.png" width=250 alt="Open Link" align="center" %}

{% include_relative diagram.htmlsnippet %}

## Legend

<img src="legend/init-box.svg" alt="Initialize"> Represents the entry point of the program. This could be a constructor of the class, connection method, macro entry point

<img src="legend/interface-box.svg" alt="Interface Box"> Represents the interface (object) of the SOLIDWORKS API

<img src="legend/cast.svg" alt="Casting"> Represents the relation between interfaces via direct casting (Query Interface)

<img src="legend/method-property.svg" alt="Method or Property"> Represents the relation between interface via method or property

<img src="legend/group.svg" alt="Group"> Represents a group of interfaces

<img src="legend/selection-interface-box.svg" alt="Selection Interface Box"> Represents an interface which is a selectable object in SOLIDWORKS User Interface

<img src="legend/selection-box.svg" alt="Selection Box"> Represents the placeholder for all selectable objects (i.e. the interfaces with blue background)

<img src="legend/etc-box.svg" alt="Etc. Box"> Represents the placeholder for other interface which are part of this group

<img src="legend/etc-box-wildcard.svg" alt="Etc. Box with wildcard"> Represents the placeholder for other interfaces which are part of this group. Interfaces name will match the wildcard
