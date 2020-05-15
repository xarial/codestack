---
layout: article
title: Automation of SOLIDWORKS using SOLIDWORKS API in vbScript
caption: vbScript
description: Introduction to automation of SOLIDWORKS using API with vbScript
---
vbScript is a popular scripting language based on Visual Basic. It is lightweight and natively supported by Windows. The code can be edited in any text editor (e.g. Notepad).

Script can be run by executing it directly (i.e. double click) or from command line. Command line option also supports input arguments.

vbScript is late bound and doesn't require the explicit declaration of variable type using the *As* keyword.

vbScript supports creation or connection to COM objects via ::CreateObject and ::GetObject methods which means that it can use SOLIDWORKS API for automation.

Use the following line to connect to the instance of SOLIDWORKS

~~~ vb
Dim swApp
Set swApp = CreateObject("SldWorks.Application")
~~~