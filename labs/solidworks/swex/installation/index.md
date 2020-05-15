---
layout: article
title: Installation and updates of SwEx.AddIn Framework for SOLIDWORKS add-ins
caption: Installation And Updates
description: Instructions on installing and updating the SwEx.AddIn framework for developing SOLIDWORKS add-ins in C# and VB.NET
image: swex-nuget-package.png
toc-group-name: labs-solidworks-swex
order: 1
---
## Installing NuGet package

Select *Manage NuGet Packages...* command from the context menu of the project in Visual Studio

![Manage NuGet Packages... command in the project context menu](manage-nuget-packages.png){ width=250 }

Search for *CodeStack.SwEx* in the search box. Once found click *Install* button for the required framework.

![CodeStack.SwEx.AddIn NuGet package](swex-nuget-package.png)

This will install all required libraries to the project.

## Preparing the project

Set the *Embed Interop Types* to *False* for the SOLIDWORKS Interop libraries as shown below.

![Disabling the option to embed interop types for SOLIDWORKS interops](sw-interops-embed-inteop-types-false.png){ width=300 }

Check the *Register for COM Interop* option in project properties:

For C# project this option can be found in *Build* tab:

![Register for COM Interop option in C# project](register-for-com-interops-csharp.png){ width=300 }

For VB.NET project this option can be found in *Compile* tab:

![Register for COM Interop option in VB.NET project](register-for-com-interops-vbnet.png){ width=300 }

## Updates

SwEx framework is actively developing and new features and bug fixes released very often. 

Nuget provides very simple way of upgrading the library versions. Simply navigate to Nuget Package manager and check for updates:

![Updating nuget packages](update-nuget-packages.png)

In order to see the release notes, follow the links below for the corresponding library.

* [SwEx.AddIn Release Notes](https://docs.codestack.net/swex/add-in/html/version-history.htm)
* [SwEx.PMPage Release Notes](https://docs.codestack.net/swex/pmpage/html/version-history.htm)
* [SwEx.MacroFeature Release Notes](https://docs.codestack.net/swex/macro-feature/html/version-history.htm)

In some cases updating the libraries may reset the *Embed Interop Types* option to *True* for SOLIDWORKS interop assemblies.

![SOLIDWORKS interop option is reset to True after the update](embed-interop-true.png){ width=350 }

It is recommended to set it back to *False*.

## Supporting multiple versions of the SwEx framework

Methods signatures and behaviour of SwEx framework might change in new versions. SwEx libraries are strong named which prevents the compatibility conflict in case several add-ins loaded in the same session of SOLIDWORKS referencing different versions of framework.