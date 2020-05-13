---
layout: article
title: Installing the SOLIDWORKS add-in by creating the msi-installer
caption: Installer
description: Creating the installer .msi package for deploying applications for SOLIDWORKS
lang: en
image: /images/codestack-snippet.png
---
Installer package (.msi) is the most robust way to deliver the best user experience when deploying the application. Installers can provide friendly step-by-step wizard with ability to specify options while installing the products. There are multiple installer options available

[Microsoft Visual Studio Installer Projects](vsi) would provide the easiest and quickest way to create an installer from the built binaries. This option however has a limited functionality and flexibility when customizing the installer.

[WiX](wix) is a popular free framework for creating the installers by defining the rules in XML format. This framework provides extensive flexibility and allows to build any customization into the installer.

Another options include but not limited to

* [InstallShield](https://en.wikipedia.org/wiki/InstallShield)
* [Nullsoft Scriptable Install System](https://en.wikipedia.org/wiki/Nullsoft_Scriptable_Install_System)
* [Orca](https://docs.microsoft.com/en-us/windows/desktop/msi/orca-exe)