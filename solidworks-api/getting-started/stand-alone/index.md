---
layout: article
title: How to create stand-alone (exe) applications using SOLIDWORKS API
caption: Stand-Alone Application
description: Two approaches to connect to SOLIDWORKS instance from the COM-compatible programming languages
lang: en
image: /solidworks-api/getting-started/stand-alone/reg-edit-clsid_prog_id.png
labels: [article, clsid, instance, out-of-process, progid, rot, sdk, solidworks api, stand-alone]
redirect_from:
  - /2018/03/connect-to-solidworks-from-stand-alone.html
order: 4
---
In this article I will discuss 2 generic approaches connecting to SOLIDWORKS instance from the COM-compatible programming languages (e.g. C#, VB.NET, C++, Visual Basic 6) in order to utilize SOLIDWORKS API.  

This is optional detailed explanation of these approaches.
Please follow the links below to access articles which demonstrate how to create a sample project and connect to SOLIDWORKS instance:  

* [Using C#]({{ "/solidworks-api/getting-started/stand-alone/connect-csharp" | relative_url }})
* [Using VB.NET]({{ "/solidworks-api/getting-started/stand-alone/connect-vbnet" | relative_url }})
* [Using C++]({{ "/solidworks-api/getting-started/stand-alone/connect-cpp" | relative_url }})

## Method A - Activator and ProgId
### Connecting by creating an instance via **Prog**ram **Id**entified (progid) or Global Unique COM **Cl**a**s**s **Id**entifier (CLSID)

There are 2 type of program identifiers for SOLIDWORKS: version independent and version specific.  

Program identifiers are registered in the Windows Registry:  

{% include img.html src="reg-edit-clsid.png" width=640 alt="Class Id in the Windows registry" align="center" %}

In the example above program identifier of the **SldWorks.Application.23** corresponds to the COM class identifier **{D66FBAAE-4150-402F-8581-75D1652D696A}**  

More information about this object (like type library class identifier, COM server location [i.e. path to **sldworks.exe**]) can be found at the registry branch related to the class identifier (i.e. **HKEY_CLASSES_ROOT\CLSID\{D66FBAAE-4150-402F-8581-75D1652D696A}**)  

{% include img.html src="reg-edit-clsid_prog_id.png" width=640 alt="Prog Id in the Windows registry" align="center" %}

Version independent program identifier will be identical for all versions of SOLIDWORKS and equal to **"SldWorks.Application"**.

If you use version independent identifier this will ensure that your code will be valid for any environment where SOLIDWORKS is installed.
This would however introduce ambiguity where multiple versions of SOLIDWORKS are installed.
In this case your program will connect to the version last installed or modified in the computer.

To use version specific program identifier it is required to specify the revision number after the program identifier, i.e. **"SldWorks.Application.RevisionNumber"**.
Please refer the table below for the list of SOLIDWORKS versions and its revision numbers:

**Version**|**Revision**
SOLIDWORKS 2005|13
SOLIDWORKS 2006|14
SOLIDWORKS 2007|15
SOLIDWORKS 2008|16
SOLIDWORKS 2009|17
SOLIDWORKS 2010|18
SOLIDWORKS 2011|19
SOLIDWORKS 2012|20
SOLIDWORKS 2013|21
SOLIDWORKS 2014|22
SOLIDWORKS 2015|23
SOLIDWORKS 2016|24
SOLIDWORKS 2017|25
SOLIDWORKS 2018|26
SOLIDWORKS 2019|27
SOLIDWORKS 2020|28

It is possible to get the revision number of SOLIDWORKS session via [ISldWorks::RevisionNumber](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~revisionnumber.html) method.
The returned value is a string in the format: **25.1.0** where first number is a revision number.  

There are few limitations when using this method:  

* It is not always predictable whether this method will connect to already running instance of SOLIDWORKS or will create new one
* It is not possible to specify which of the running SOLIDWORKS sessions to connect to (e.g. when more than one SOLIDWORKS session is open)
* If new session is created as the result of running this method this session will be invisible by default and started with */embed* flag.
That means that session is started lightweight and no add-ins are loaded.
This was designed to allow embedding OLE objects into the 3rd party applications (such as Microsoft Office).

{% include img.html src="excel-ole-object.png" width=400 alt="SOLIDWORKS Part Document OLE object in Excel" align="center" %}

* It is not possible to create more than one active sessions of SOLIDWORKS

## Method B - Running Object Table (ROT)

### Connecting by querying the COM instance from the **R**unning **O**bject **T**able (ROT)

When COM server creates an object instance it creates a moniker for this instance and registers it in the Running Objects Table (ROT).
ROT enables interprocess communication with 3rd party applications by allowing to lookup the objects from the running processes via Windows APIs ([GetRunningObjectTable](https://msdn.microsoft.com/en-us/library/windows/desktop/ms684004(v=vs.85).aspx))

Below is an example of Running Object Table with several registered COM objects:  

>!{00024505-0014-0000-C000-000000000046}

>!Microsoft Visual Studio Telemetry:11004

>!{31F45B04-7198-45ED-A13F-F224A4A1686A}

>SolidWorks_PID_15212

>!VisualStudio.DTE.14.0:16144

* Using this approach it is possible to connect to any session of SOLIDWORKS from its process id
* It is possible to create as many sessions as needed by starting new SOLIDWORKS instance via shell or start process APIs

> Object might not be successfully retrieved form the ROT if the SOLIDWORKS application and the stand-alone application are run with different permission levels (e.g. one is run as administrator while other is not). Run them under the same user to enable communication.

Please follow the links at the beginning of the articles for the detailed guides with code examples for connecting to SOLIDWORKS instance.
