---
title: Life cycle of SOLIDWORKS macro feature
caption: Lifecycle
description: Explanation of the SOLIDWORKS macro feature behavior and life cycle
toc-group-name: labs-solidworks-swex
order: 2
---
Macro feature resides in the model and saved together with the document. Macro feature can handle various events during its lifecycle

* Regeneration. Override [OnRebuild](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_OnRebuild.htm) method to handle this event.
* Editing. Override [OnEditDefinition](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_OnEditDefinition.htm) method to handle this event.
* Updating state. Override [OnUpdateState](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_OnUpdateState.htm) method to handle this event.

Macro feature is a singleton service. Do not create any class level variables in the macro feature class. If it is required to track the lifecycle of particular macro feature use
[Feature Handler](feature-handler)