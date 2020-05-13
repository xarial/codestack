---
layout: article
title: Installation and updates of SwEx.AddIn Framework for SOLIDWORKS add-ins
caption: Installation And Updates
description: Instructions on installing and updating the SwEx.AddIn framework for developing SOLIDWORKS add-ins in C# and VB.NET
lang: en
image: /labs/solidworks/swex/installation/swex-nuget-package.png
toc_group_name: labs-solidworks-swex
order: 1
---
## Installing NuGet package

Select *Manage NuGet Packages...* command from the context menu of the project in Visual Studio

{% include img.html src="manage-nuget-packages.png" height=250 alt="Manage NuGet Packages... command in the project context menu" align="center" %}

Search for *CodeStack.SwEx* in the search box. Once found click *Install* button for the required framework.

{% include img.html src="swex-nuget-package.png" alt="CodeStack.SwEx.AddIn NuGet package" align="center" %}

This will install all required libraries to the project.

## Preparing the project

Set the *Embed Interop Types* to *False* for the SOLIDWORKS Interop libraries as shown below.

{% include img.html src="sw-interops-embed-inteop-types-false.png" height=300 alt="Disabling the option to embed interop types for SOLIDWORKS interops" align="center" %}

Check the *Register for COM Interop* option in project properties:

For C# project this option can be found in *Build* tab:

{% include img.html src="register-for-com-interops-csharp.png" height=300 alt="Register for COM Interop option in C# project" align="center" %}

For VB.NET project this option can be found in *Compile* tab:

{% include img.html src="register-for-com-interops-vbnet.png" height=300 alt="Register for COM Interop option in VB.NET project" align="center" %}

## Updates

SwEx framework is actively developing and new features and bug fixes released very often. 

Nuget provides very simple way of upgrading the library versions. Simply navigate to Nuget Package manager and check for updates:

{% include img.html src="update-nuget-packages.png" alt="Updating nuget packages" align="center" %}

In order to see the release notes, follow the links below for the corresponding library.

* [SwEx.AddIn Release Notes](https://docs.codestack.net/swex/add-in/html/version-history.htm)
* [SwEx.PMPage Release Notes](https://docs.codestack.net/swex/pmpage/html/version-history.htm)
* [SwEx.MacroFeature Release Notes](https://docs.codestack.net/swex/macro-feature/html/version-history.htm)

In some cases updating the libraries may reset the *Embed Interop Types* option to *True* for SOLIDWORKS interop assemblies.

{% include img.html src="embed-interop-true.png" width=350 alt="SOLIDWORKS interop option is reset to True after the update" align="center" %}

It is recommended to set it back to *False*.

## Supporting multiple versions of the SwEx framework

Methods signatures and behaviour of SwEx framework might change in new versions. SwEx libraries are strong named which prevents the compatibility conflict in case several add-ins loaded in the same session of SOLIDWORKS referencing different versions of framework.