---
layout: article
title: Block model editing using SOLIDWORKS API
caption: Block Model Editing
description: Example demonstrate different ways of disabling the model editing
labels: [block editing, block model, example, lock, menu, solidworks api]
redirect-from:
  - /2018/03/block-model-editing.html
---
This example demonstrate different ways of disabling the model editing from SOLIDWORKS API: 

* Blocking menu - user is not able to invoke menu commands. This feature is usually used when property manager page is displayed and there should be no commands invoked
* Blocking model editing - model is a view only and cannot be changed
* Full block - editing and view manipulations are disabled

It is required to debug macro step-by-step to see the different SOLIDWORKS API functions in action.

{% code-snippet { file-name: Macro.vba } %}