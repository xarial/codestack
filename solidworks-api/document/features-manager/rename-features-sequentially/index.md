---
layout: sw-tool
title: SOLIDWORKS macro renames all features in model sequentially
caption: Rename All Features Sequentially
description: Macro renames all the features in the order preserving the base names using SOLIDWORKS API
image: /solidworks-api/document/features-manager/rename-features-sequentially/rename-features-sequentially.png
logo: /solidworks-api/document/features-manager/rename-features-sequentially/sequntial-features.svg
labels: [feature, rename]
categories: sw-tools
group: Model
---
{% include img.html src="rename-features-sequentially.png" alt="Features renamed sequentially" align="center" %}

This macro renames all the features in active model in the sequential order using SOLIDWORKS API, preserving the base names .

Only indices are renamed and the base name is preserved. For example *Sketch21* will be renamed to *Sketch1* for the first appearance of the sketch feature.

## Notes

* Only features with number at the end will be renamed (e.g. *Front Plane* will not be renamed to *Front Plane1* and *My1Feature* will not be renamed)
* Case is ignored (case insensitive search)
* Only modelling features are renamed (the ones created after the Origin feature)

{% code-snippet { file-name: Macro.vba } %}
