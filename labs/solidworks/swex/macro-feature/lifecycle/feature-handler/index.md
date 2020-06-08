---
title: Handling the life cycle of SOLIDWORKS macro feature
caption: Feature Handler
description: Using SOLIDWORKS macro feature handler to manage the life cycle of the macro feature in SwEx.MacroFeature framework
toc-group-name: labs-solidworks-swex
order: 4
---
[MacroFeatureEx{TParams, THandler} Class](https://docs.codestack.net/swex/macro-feature/html/T_CodeStack_SwEx_MacroFeature_MacroFeatureEx_2.htm) overload of macro feature allows defining the handler class which will be created for each feature. This provides a simple way to track the macro feature lifecycle (i.e. creation time and deletion time).

{% code-snippet { file-name: FeatureHandler.cs } %}

Instance of the handler class will be created and disposed by framework. This approach is useful when macro feature needs to monitor the events of a specific file it resides.