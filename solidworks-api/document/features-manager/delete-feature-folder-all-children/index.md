---
layout: sw-tool
title: SOLIDWORKS Macro to delete feature folder with all children features
caption: Delete Feature Folder With All Children Features
description: Macro allows to delete all of the features in the selected folder in one click using SOLIDWORKS API
image: /solidworks-api/document/features-manager/delete-feature-folder-all-children/deleted-folder-features.png
logo: /solidworks-api/document/features-manager/delete-feature-folder-all-children/deleted-folder-features.svg
labels: [delete folder, feature manager, folder, solidworks api, utility]
categories: sw-tools
group: Model
redirect-from:
  - /2018/04/solidworks-api-feature-manager-delete-feature-folder-with-all-children.html
---
When deleting the top folder in SOLIDWORKS features tree all sub features are not get deleted so it is required to select all of them one-by-one in order to delete folder content.

This is not always possible to do in one step due to the features relations:  

![Manually deleting the folder feature](delete-features-manually.gif){ width=400 }

The macro below allows to delete all of the features in the selected folder in one click using SOLIDWORKS API. Nested folders are also supported.

![Deleting the folder with all children features](delete-folder-with-features.png){ width=400 }

Macro can optionally display the confirmation dialog with the list of features about to be deleted

{% code-snippet { file-name: Macro.vba } %}
