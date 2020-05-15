---
layout: article
title: Create C++ Stand-Alone (exe) application for SOLIDWORKS
caption: Create C++ Stand-Alone Application for SOLIDWORKS
description: Guide for how to connect to SOLIDWORKS application from out-of-process (a.k.a Stand-Alone) application (e.g. MFC, Win32 Console Application) using C++ and Microsoft Visual Studio
order: 3
image: proj-templ.png
labels: [c++, CoCreateInstance, create instance, example, getobject, rot, sdk, solidworks api, tlb, type library]
redirect-from:
  - /2018/03/create-c-stand-alone-application-for_5.html
---
In this tutorial I will demonstrate how to connect to SOLIDWORKS application from out-of-process (a.k.a Stand-Alone) application (e.g. MFC, Win32 Console Application) using C++ and Microsoft Visual Studio.

For more detailed explanation of the approaches discussed in this article refer the [Connect To SOLIDWORKS From Stand-Alone Application](/solidworks-api/getting-started/stand-alone/) article.

## Creating new project

I will be using Microsoft Visual Studio development environment. You can use any edition of Visual Studio.
The same code will work in Professional, Express or Community editions. Follow this link to download [Visual Studio](https://www.visualstudio.com/vs/community/)

* Open Visual Studio
* Start new project:

![Creating new project in Visual Studio](new-project.png){ width=400 }

* Select the project template. I would recommend to start with Win32 Console Application project template as it contains the minimum pregenerated code:

![Selecting the Win32 Console Application C++ project template](proj-templ.png){ width=640 }

* Check the ATL option in the project wizard

![Win32 Console Application template settings](apps-settings.png){ width=640 }

* Link directory where SOLIDWORKS type libraries are located.
This is an installation directory of SOLIDWORKS (Go to Project Properties, select C/C++ and browse the path in the *Additional Include Directories* field):

![Additional Include Directories option in C++ project](add-incl-dir.png){ width=640 }

Now we can add the code to connect to SOLIDWORKS instance.  

## Creating or connecting to instance

Probably the most common and quick way to connect to COM server is using the [CoCreateInstance](https://msdn.microsoft.com/en-us/library/windows/desktop/ms686615(v=vs.85).aspx) function.  

{% code-snippet { file-name: CreateGetInstance.cpp } %}

## Getting the running instance via ROT

In order to connect to already running specific session of SOLIDWORKS or to be able to create multiple sessions you can use Running Object Table APIs.
Please read the [Connect To SOLIDWORKS From Stand-Alone Application](/solidworks-api/getting-started/stand-alone#method-b---running-object-table-rot) article for more details about this approach.

{% code-snippet { file-name: CreateInstance.cpp } %}

In the above example new session of SOLIDWORKS is launched by starting new process from SOLIDWORKS application installation path.
*ConnectToSwApp* function requires the full path to **sldworks.exe** as first parameter and timeout in seconds as second parameter.
Timeout will ensure that the application won't be locked in case process failed to start.
