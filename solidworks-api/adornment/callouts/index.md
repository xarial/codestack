---
layout: article
title: Using Callouts object for model annotation in SOLIDWORKS API
caption: Callouts
description: Using Callouts for annotating models (similar to balloons), linking to the entities and displaying custom data using SOLIDWORKS API
lang: en
order: 1
image: /images/codestack-snippet.png
labels: [callout, balloons]
---
Callouts in SOLIDWORKS are balloon like objects which can be attached to the entities (usually with selection) and display additional information about the entity. Callouts are not scaled with the model and stay in the same orientation even when model rotates.

Callouts are temp objects and usually destroyed after selection is cleared or operation complete.

The most common example of callouts in SOLIDWORKS is a Measure tool. When entities are selected the measurement results are displayed in the callouts.

SOLIDWORKS API enables the ability for creating callouts via the [ISwCalloutHandler Interface](http://help.solidworks.com/2018/english/api/swpublishedapi/solidworks.interop.swpublished~solidworks.interop.swpublished.iswcallouthandler.html) interface. THis handler allows to create a definition of the callout and handle related events.

This section contains macros and code examples utilizing SOLIDWORKS API to create, display and handle events of the callouts.