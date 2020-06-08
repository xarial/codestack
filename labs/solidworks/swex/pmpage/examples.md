---
title: Examples and source code for SwEx.PMPage framework
caption: Examples
description: Collection of examples using the SwEx.PMPage framework for SOLIDWORKS
toc-group-name: labs-solidworks-swex
order: 4
---
This section contains examples which are using SwEx.PMPage framework and SOLIDWORKS API for developing advanced Property Manager Pages in SOLIDWORKS add-ins.

## Dummy add-in used to debug main framework functionality (C#)
[Source Code](https://github.com/codestackdev/swex-pmpage/tree/master/Samples/AddIn)

This project contains short example which is utilizing all functionality of SwEx.PMPage framework, although this project doesn't perform any useful operations it is a good starting point to explore SwEx.PMPage API.

## Insert Note Example (C#)
* [Source Code](https://github.com/codestackdev/swex-examples/tree/master/pmpage/InsertNote/csharp)

This examples is using Property Manager page to collect the user input required for insertion of the repetitive note in the SOLIDWORKS drawing document:

* Note text
* Note position
* Optional entity to attach

Use inputs are stored and can be reused in different SOLIDWORKS sessions