---
layout: article
title: Create C# stand-alone application for SOLIDWORKS API automation
caption: Create C# Stand-Alone Application for SOLIDWORKS
description: Guide of how to connect to SOLIDWORKS application from out-of-process (a.k.a Stand-Alone) application (e.g. Windows Forms, Windows Console) using C# and Microsoft Visual Studio
order: 1
image: /solidworks-api/getting-started/stand-alone/connect-csharp/proj-template.png
labels: [activator, c#, create instance, example, getobject, rot, sdk, solidworks api]
redirect_from:
  - /2018/03/create-c-stand-alone-application-for.html
---
In this tutorial I will demonstrate how to connect to SOLIDWORKS application from out-of-process (a.k.a Stand-Alone) application (e.g. Windows Forms, Windows Console) using C# and Microsoft Visual Studio.  

For more detailed explanation of the approaches discussed in this article please read the [Connect To SOLIDWORKS From Stand-Alone Application](/solidworks-api/getting-started/stand-alone/) article.

## Creating new project

I will be using Microsoft Visual Studio development environment. You can use any edition of Visual Studio. The same code will work in Professional, Express or Community editions. Follow this link to download [Visual Studio](https://www.visualstudio.com/vs/community/)  

1. Open Visual Studio. 
1. Start new project:

{% include img.html src="new-project.png" width=400 alt="Creating new project in Visual Studio" align="center" %}

1. Select the project template. I would recommend to start with Console Application project template as it contains the minimum pregenerated code:

{% include img.html src="proj-template.png" width=640 alt="Selecting C# Console Application project template" align="center" %}

1. Add reference to SolidWorks Interop library. Interop libraries are located at **SOLIDWORKS Installation Folder**\api\redist\SolidWorks.Interop.sldworks.dll for projects targeting Framework 4.0 onwards and **SOLIDWORKS Installation Folder**\api\redist\CLR2\SolidWorks.Interop.sldworks.dll for projects targeting Framework 2.0 and 3.5.

{% include img.html src="add-ref.png" width=320 alt="Adding assembly references to the project" align="center" %}

For projects targeting Framework 4.0 I recommend to set the **[Embed Interop Types](https://docs.microsoft.com/en-us/dotnet/framework/interop/type-equivalence-and-embedded-interop-types)** option to false.
Otherwise it is possible to have unpredictable behaviour of the application when calling the SOLIDWORKS API due to a type cast issue.  

{% include img.html src="embed-interop-types.png" height=320 alt="Option to embed interop assemblies" align="center" %}

Now we can add the code to connect to SOLIDWORKS instance.  

## Creating or connecting to instance

Probably the most common and quick way to connect to COM server is using the [Activator::CreateInstance](https://msdn.microsoft.com/en-us/library/system.activator.createinstance(v=vs.110).aspx) method.  

{% include_relative CreateGetInstance.cs.codesnippet %}

This method will construct the instance of the type from the type definition. As SOLIDWORKS application is registered as COM server we can create the type from its program identifier via [Type::GetTypeFromProgID](https://msdn.microsoft.com/en-us/library/system.type.gettypefromprogid(v=vs.110).aspx) method.
Please read the [Connect To SOLIDWORKS From Stand-Alone Application](/solidworks-api/getting-started/stand-alone#method-a---activator-and-progid) article for explanations of limitation of this approach.  

Alternatively you can connect to active (already started) session of SOLIDWORKS using the [Marshal::GetActiveObject](https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.getactiveobject(v=vs.110).aspx) method.
This approach will ensure that  there will be no new instances of SOLIDWORKS created and will throw an exception if there is no running SOLIDWORKS session to connect to.

{% include_relative GetInstance.cs.codesnippet %}

## Getting the running instance via ROT

In order to connect to already running specific session of SOLIDWORKS or to be able to create multiple sessions you can use Running Object Table APIs.
Please read the [Connect To SOLIDWORKS From Stand-Alone Application](/solidworks-api/getting-started/stand-alone#method-b---running-object-table-rot) article for more details about this approach.

{% include_relative CreateInstance.cs.codesnippet %}

In the above example new session of SOLIDWORKS is launched by starting new process from SOLIDWORKS application installation path.
*StartSwApp* function requires the full path to **sldworks.exe** as first parameter and optional timeout in seconds as second parameter.
Timeout will ensure that the application won't be locked in case process failed to start.  

You can also make this call asynchronous and display some progress indication in your application while SOLIDWORKS process is starting:

{% include_relative StartSwAppAsync.cs.codesnippet %}
