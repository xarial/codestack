---
layout: article
title: VBA macro to hide all selected features from the SOLIDWORKS file tree
caption: Hide Features In The Tree
description: VBA macro which hides features and makes them invisible in the SOLIDWORKS Feature Manager tree
image: /solidworks-api/document/features-manager/hide-features/hidden-features.png
labels: [feature,hide,invisible]
---
This VBA macro allows to make invisible selected features in the tree. The features still continue to be fully operational and visible in the graphics area (e.g. planes), but not visible in the feature manager tree.

Even default features (such as planes) can be made invisible.

![Sketch, Right and Top planes hidden in the feature manager tree](hidden-features.png)

To show the hidden features use the [Reveal Hidden Features](/solidworks-api/document/features-manager/reveal-hidden-features/) macro.

{% code-snippet { file-name: Macro.vba } %}
