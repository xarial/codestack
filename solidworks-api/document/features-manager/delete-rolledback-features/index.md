---
layout: sw-tool
caption: Delete rolled back features
title: Macro to delete all features which are in the rolled back state in SOLIDWORKS document
description: VBA macro finds and deletes all features below the rollback bar
image: rollback-feature.svg
group: Model
---
![Features rolled back in the feature manager tree](rolled-back-features.png)

This VBA macro deletes all features which are below the rollback bar.

{% code-snippet { file-name: Macro.vba } %}