---
layout: article
caption: Insert Location Label
title: Add location label to a drawing view
description: VBA macro which demonstrates how to add location label to a drawing view
image: location-label.png
---
![Inserting location label](location-label.png)

This VBA macro provides a workaround for missing SOLIDWORKS API to insert the location label to a drawing view.

Specify the name of the view as **VIEW_NAME** constant.

> Only views compatible with location label are supported, e.g. auxillary, detailed, etc.

{% code-snippet { file-name: Macro.vba } %}