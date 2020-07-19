---
caption: Create MultiBoss-Extrude Macro Feature
title: Create MultiBoss-Extrude VBA macro feature using SOLIDWORKS API
description: VBA example demonstrates how to create parametric macro feature to create extrude from multiple sketches with an editing and preview ability
image: multiboss-extrude.png
---
![Property Manager Page and preview for MultiBoss-Extrude Macro Feature](multiboss-extrude.png) { width=500 }

This VBA macro demonstrates how to create parametric SOLIDWORKS macro feature to create single extrude from multiple sketches using VBA.

Watch video below which demonstrates the macro and explains how macro was built and how it works.

{% youtube id: EAx78xOsU3s %}

Create the following macro structure and copy the code snippets to the corresponding modules and classes.

![Macro modules and classes](macro-project-structure.png)

## Macro Module

Entry point of the macro. Use this to insert new macro feature.

{% code-snippet { file-name: Macro.vba } %}

## Geometry Module

Module contains helper functions for building temp geometry of extrudes from input sketches

{% code-snippet { file-name: Geometry.vba } %}

## MacroFeature Module

Implements the behavior of macro feature: regeneration and editing

{% code-snippet { file-name: MacroFeature.vba } %}

## PropertyPage Class Module

Implements the property manager page interface for the macro feature.

{% code-snippet { file-name: PropertyPage.vba } %}

## Controller Class Module

Connects the property manager page inputs to corresponding functionality (i.e. Edit or Insert)

{% code-snippet { file-name: Controller.vba } %}

[Download sample model](MacroFeatureMultiExtrude.SLDPRT)