---
layout: article
title: Hosting eDrawings control in Windows Forms
caption: Hosting Control in WinForms
description: Guide of hosting eDrawings control in C# Widndows Forms
image: /edrawings-api/gettings-started/winforms/edrawings-winform.png
labels: [edrawings,host,winforms]
---
This tutorial is a step-by-step guide of hosting the eDrawings control in the Windows Forms using C#.

## Creating new Windows Form project

* Open Visual Studio
* Create new project and select *Windows Forms Application* in the *Visual C#* templates section

![New Windows Forms project](visualstudio-new-project.png){ width=550 }

## Adding eDrawings interop

It is required to add interop for eDrawings control. Note that interop only contains the signature of the methods and classes (it doesn't contain implementation). eDrawings stil needs to be installed to use its API. However SOLIDWORKS application doesn't need to be installed.

Locate the eDrawings interop library by searching for *eDrawings.Interop.EModelViewControl.dll* file in the installation folder of eDrawings. Usually the path will be equal to 
*C:\Program Files\Common Files\eDrawings[Version]\eDrawings.Interop.EModelViewControl.dll*

I would recommend to disable the embedding of the interop to avoid potential cast errors:

![Disable the embedding of interop files](embed-edrawings-interops.png){ width=350 }

## Creating eDrawings control wrapper

By default eDrawings API doesn't provide the .NET control to be hosted on the Windows Forms. So it is required to create the corresponding wrapper by implementing the [AxHost](https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.axhost) control and providing eDrawings GUID in the constructor. In the below implementation the version independent (*22945A69-1191-4DCF-9E6F-409BDE94D101*) guid is used which means that you application is backward compatible and will work with newer versions of eDrawings. Find the version dependent guids for eDrawing controls in the Registry if required.

### eDrawingHost.cs

{% code-snippet { file-name: eDrawingHost.cs } %}

> eDrawings control is not loaded immediately and calling the [AxHost::GetOcx](https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.axhost.getocx) directly after the constructor will result in null reference. Calling this method when control is not fully loaded might result into the deadlock.

## Adding control to Windows Form

Add new form to the project and name it *MainForm*.

The solution tree will look similar to the one below:

![Project solution tree](solution-explorer.png){ width=350 }

Add the following code to the form code behind. Set the path to the SOLIDWORKS file as the *FILE_PATH* constant. The code will wait until eDrawings control is fully loaded and open the specified file automatically.

![File opened in the eDrawing control hosted in Windows Form](edrawings-winform.png){ width=350 }

### MainForm.cs

{% code-snippet { file-name: MainForm.cs } %}

Source code is available on [GitHub](https://github.com/codestackdev/solidworks-api-examples/tree/master/edrawings-api/eDrawingsWinFormsHost)
