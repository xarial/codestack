---
layout: article
title: Recursively Traverse Files And Folders In Vault Using SOLIDWORKS PDM API
caption: Traverse Folder Recursively
description: VBA example to traverse and list all files and folders from the selected folder in SOLIDWORKS PDM vault using SOLIDWORKS PDM API
lang: en
image: /solidworks-pdm-api/files/traverse/pdm-folder-structure-output.png
labels: [traverse,vault,browse folder]
---
This VBA example demonstrates how to traverse files and folders in the SOLIDWORKS PDM vault using SOLIDWORKS PDM API.

Macro displays the built-in folder browse dialog for the folder to traverse:

{% include img.html src="browse-folder.png" width=250 alt="Built-in PDM Folder Browse dialog" align="center" %}

Macro recursively traverses files and sub folders and outputs the file or folder name, id, level to the VBA Editor immediate window.

{% include img.html src="pdm-folder-structure-output.png" width=350 alt="Folders and files structure output to immediate window of VBA Editor" align="center" %}

This macro can traverse the tree even if it is not [cached locally]({{ "/solidworks-pdm-api/files/local-cache/" | relative_url }})

{% include_relative Macro.vba.codesnippet %}
