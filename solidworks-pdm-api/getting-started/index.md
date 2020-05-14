---
layout: article
title: 'Getting Started: SOLIDWORKS PDM API Development'
caption: Getting Started
description: Introduction to SOLIDWORKS PDM API, explanation of different ways of accessing API from add-ins and stand-alone applications
image: /images/codestack-snippet.png
labels: [pdm api,getting started]
order: 1
---
This section introduces to SOLIDWORKS PDM API. It explains the differences between PDM add-ins and stand-alone application and provides detailed guidelines of creating the ones.

The root object in SOLIDWORKS PDM API is a [IEdmVault5](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmvault5.html) which provides an access to various section of the functionality.

This interface can be explicitly cast to another manager interfaces, such as [IEdmAddInMgr9](http://help.solidworks.com/2018/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmAddInMgr9.html?id=96f8b929514a423d8cb220fbe54bb940#Pg0), [IEdmRevisionMgr3](http://help.solidworks.com/2018/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmRevisionMgr3.html?id=755088fcb7fc40a99dfb42fb5e5b237e#Pg0), etc.

The most popular way of extending the system is by implementing the add-in via [IEdmAddIn5](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmaddin5.html) SOLIDWORKS PDM API interface.

## Interops in .NET

If you are building the application in .NET (C# or VB.NET) you will need to use SOLIDWORKS PDM API interop to access the signatures of API methods.

### Framework 4.0 or newer

You need to add the reference to *EPDM.Interop.epdm.dll* which is located in the installation folder of PDM (usually *C:\Program Files\SOLIDWORKS PDM\EPDM.Interop.epdm.dll*).

Note, although you can add the reference to *EdmInterface.dll* (type library) this will generate the *Interop.EdmLib.dll* which can be used by .NET, however this interop will not have a strong name which may introduce conflicts with other add-ins.

It is recommended to set the *Embed Interop Types* option to *False* for the interop otherwise the add-in may misbehave.

### Framework 2.0 or older

Newer versions of SOLIDWORKS PDM do not provide the interop compatible with .NET Framework 2.0 or older. So it is required to generate this interop from the type library (*EdmInterface.dll*).

Either add this reference directly to your project (usually *C:\Program Files\SOLIDWORKS PDM\EdmInterface.dll*), this will generate the *Interop.EdmLib.dll* in the bin folder after rebuild which you can reference by other projects.

Or alternatively use the [tlbim.exe](https://docs.microsoft.com/en-us/dotnet/framework/tools/tlbimp-exe-type-library-importer) utility to generate the interop using the following command:

~~~
> TlbImp.exe "EdmInterface.dll" "/out:Interop.EdmLib.dll" /namespace:EdmLib
~~~