---
layout: article
title: Handle custom properties modification events (add, delete, change) using SOLIDWORKS API
caption: Handle Modification Events
description: Handling all events related to the modification of general or configuration specific custom properties from the SOLIDWORKS API. Workaround for the issue when AddCustomPropertyNotify, DeleteCustomPropertyNotify, ChangeCustomPropertyNotify events are not raised
image: /solidworks-api/data-storage/custom-properties/handle-events/custom-properties-events-console.png
labels: [custom property,notification]
---
SOLIDWORKS API provides notifications to handle the custom properties modifications (such as add, delete or change). These events (AddCustomPropertyNotify, DeleteCustomPropertyNotify, ChangeCustomPropertyNotify) are raised for parts, assemblies and drawings and support general and configuration specific custom properties. However since SOLIDWORKS 2018 these events are no longer raised for the custom properties modified by the user in the user interface and only support custom properties modified from SOLIDWORKS API.

The code example below provides a workaround for this issue and enables handling of the notifications regardless of the way custom properties were modified.

* Create console application and add the code below
* Run the console
* Modify custom properties. The modification results are output to the console window:

![Properties modification information output to the console](custom-properties-events-console.png)

## Program.cs

Entry point of the program

{% code-snippet { file-name: Program.cs } %}

## CustomPropertiesEventsHandler.cs

{% code-snippet { file-name: CustomPropertiesEventsHandler.cs } %}
