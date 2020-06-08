---
title: Examples with source code built using SwEx.MacroFeature framework for SOLIDWORKS API
caption: Examples
description: Collection of examples using the SwEx.MacroFeature framework for SOLIDWORKS API
toc-group-name: labs-solidworks-swex
order: 4
---
SwEx.MacroFeature is a framework for SOLIDWORKS API enabling the simplified development and data binding of macro features.

This section lists examples and applications based on SwEx.MacroFeature framework.

## Test Macro Feature Project

[Source Code](https://github.com/codestackdev/swex-macrofeature/tree/dev/AddInExample)
Basic example of all features in the SwEx.Framework. This example doesn't perform any useful function. Use it to explore the code for snippets of the framework usage.

## Stock Master

[Source Code](https://github.com/codestackdev/stock-fit-geometry)

Utility to automate to generate stocking fit (boundary element) for the 3D geometry Documentation. Macro feature is used to generate custom stock feature which generates the cylindrical geometry around input body based on the input parameters.

## Convert Solid To Surface

[Source Code](https://github.com/codestackdev/solidworks-api-examples/tree/master/swex/macro-feature/convert-solid-to-surface)

Example macro feature which allows to convert solid body to the surface body preserving the associativity.

## Geometry++

[Source Code](https://github.com/codestackdev/geometry-plus-plus)

Advanced commands for managing the geometry in SOLIDWORKS. All commands implemented as a dynamic macro feature which modifying the existing or adding new geometry.

## Link Geometry To External File

[Source Code](https://github.com/codestackdev/solidworks-api-examples/tree/master/swex/macro-feature/link-external-file)

Example of loading body geometry from external file (similar to Insert Part functionality)