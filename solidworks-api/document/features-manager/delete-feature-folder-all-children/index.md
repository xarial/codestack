---
layout: sw-tool
title: SOLIDWORKS Macro to delete feature folder with all children features
caption: Delete Feature Folder With All Children Features
description: Macro allows to delete all of the features in the selected folder in one click using SOLIDWORKS API
lang: en
image: /solidworks-api/document/features-manager/delete-feature-folder-all-children/deleted-folder-features.png
logo: /solidworks-api/document/features-manager/delete-feature-folder-all-children/deleted-folder-features.svg
labels: [delete folder, feature manager, folder, solidworks api, utility]
categories: sw-tools
group: Model
redirect_from:
  - /2018/04/solidworks-api-feature-manager-delete-feature-folder-with-all-children.html
---
When deleting the top folder in SOLIDWORKS features tree all sub features are not get deleted so it is required to select all of them one-by-one in order to delete folder content.

This is not always possible to do in one step due to the features relations:  

{% include img.html src="delete-features-manually.gif" width=400 alt="Manually deleting the folder feature" align="center" %}

The macro below allows to delete all of the features in the selected folder in one click using SOLIDWORKS API. Nested folders are also supported.

{% include img.html src="delete-folder-with-features.png" width=400 alt="Deleting the folder with all children features" align="center" %}

Macro can optionally display the confirmation dialog with the list of features about to be deleted

{% include_relative Macro.vba.codesnippet %}
