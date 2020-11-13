---
layout: sw-tool
title: Remove all mates and fix components in SOLIDWORKS assembly
caption: Remove All Mates And Fix Components
description: VBA macro to remove all mates and fix all top level components in the active SOLIDWORKS assembly
image: remove-mates.svg
labels: [mates, remove, fix]
group: Assembly
---
![Mates in the Feature Manager Tree](feature-tree-mates.png)

This VBA macro remove all mates from the active assembly and fixes all the top level components.

This allows to significantly improve the performance of the assembly.

{% code-snippet { file-name: Macro.vba } %}
