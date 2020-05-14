---
layout: article
title: SOLIDWORKS VBA macro to compose flat BOM table using API
caption: Compose Flat Bill Of Materials (BOM)
description: Example demonstrates how to compose bill of materials from the assembly tree using SOLIDWORKS API
image: /solidworks-api/document/assembly/compose-flat-bom/bill-of-materials.png
labels: [bom, flat, top level]
---
{% include img.html src="bill-of-materials.png" width=250 alt="Bill Of Materials" align="center" %}

This example demonstrates how to compose flat (top level only) Bill Of Materials table from the assembly tree using SOLIDWORKS API.

Bill Of Materials position includes the following columns:

* Model Path
* Model Configuration
* Description (custom property)
* Price (custom property)
* Quantity (calculated)

The composed BOM is output to the immediate window of VBA editor:

{% include img.html src="flat-bom-print.png" width=250 alt="BOM Table printed in the immediate window" align="center" %}

It is not required to have a BOM Table inserted for this macro to work.

{% include_relative Macro.vba.codesnippet %}
