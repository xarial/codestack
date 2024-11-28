---
caption: Remove Flat Pattern Configurations
title: Remove flat pattern configurations (SM-FLAT-PATTERN) from SOLIDWORKS parts
description: VBA macro deletes derived SM-FLAT-PATTERN configurations of the sheet metal SOLIDWORKS parts using SOLIDWORKS API
image: 
macro-plus: vba
---

This VBA macro deletes all **\<ConfName\>SM-FLAT-PATTERN** configurations from SOLIDWORKS part file

This configuration is created automatically when flat pattern drawing view is created for the sheet metal parts. In some cases this configuration may produce incorrect flat pattern geometry (e.g. missing the unbending). IN order to fix the issue it might be required to remove this configuration and recreate a drawing view.

{% code-snippet { file-name: Macro.vba } %}