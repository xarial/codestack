---
layout: sw-tool
title: SOLIDWORKS macro renames all features in model sequentially
caption: Rename All Features Sequentially
description: Macro renames all the features in the order preserving the base names using SOLIDWORKS API
image: sequntial-features.svg
labels: [feature, rename]
group: Model
---
![Features renamed sequentially](rename-features-sequentially.png)

This macro renames all the features in active model in the sequential order using SOLIDWORKS API, preserving the base names .

Only indices are renamed and the base name is preserved. For example *Sketch21* will be renamed to *Sketch1* for the first appearance of the sketch feature.

## Notes

* Only features with number at the end will be renamed (e.g. *Front Plane* will not be renamed to *Front Plane1* and *My1Feature* will not be renamed)
* Case is ignored (case insensitive search)
* Only modelling features are renamed (the ones created after the Origin feature)

Watch [video demonstration](https://youtu.be/jsjN8zNRTuc?t=139)

{% code-snippet { file-name: Macro.vba } %}
