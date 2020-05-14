---
layout: article
title: In-Process invoking of SOLIDWORKS add-in API from out-of-process applications
caption: In-Process Invoking From Out-Of-Process Applications
description: Framework for calling the add-in API in-process form stand-alone applications or macros to gain maximum performance
image: /solidworks-api/getting-started/inter-process-communication/invoke-add-in-functions/in-process-invoking/macro-solution-tree.png
labels: [add-in api,async,performance,in-process]
order: 4
---
One of the main limitations of the stand-alone automation of COM based application automation (including SOLIDWORKS) is performance.

When hundreds of API calls need to be called from out-of-process applications, the performance may be dropped in hundreds or even thousands of times compared to in-process invocation.

The exact same limitation would apply when invoking add-in API in any of the following approaches: [via add-in object]({{ "/solidworks-api/getting-started/inter-process-communication/invoke-add-in-functions/via-add-in-object/" | relative_url }}), [via Running Object Table]({{ "/solidworks-api/getting-started/inter-process-communication/invoke-add-in-functions/via-rot/" | relative_url }}), etc.

It can be mistakenly assumed that all of the SOLIDWORKS API calls inside the add-in are invoked in-process as only single API function is called form stand-alone. But in fact all of the SOLIDWORKS API calls within the SOLIDWORKS add-in are invoked as out-of-process calls. This means that calling the add-in API would result in the same performance loses as calling the stand-alone application.

There is however a way to maximize this performance and gain the same results as in-process calls by calling this from out-of-process application.

The following add-in example implements a function to index all faces of the active assembly documents.

Add-in is developed using the [SwEx.AddIn Framework]({{ "/labs/solidworks/swex/add-in/" | relative_url }}), but the same technique can apply to add-in built with different methods.

It traverses all components, all bodies and all faces and outputs some information about the face in the trace window.

Add-in has a menu command allowing to invoke its function in-process.

{% include img.html src="face-indexer-menu.png" width=350 alt="Add-in menu to index faces" align="center" %}

Once completed the message box with the result is displayed.

{% include img.html src="add-in-result.png" width=300 alt="Result from calling the add-in command" align="center" %}

## FaceIndexer Add-In
This is a main project which implements SOLIDWORKS add-in and API object interface.

### FaceIndexerAddIn.cs

Add-in class

{% code-snippet { file-name: face-indexer/FaceIndexerAddIn.cs } %}

### FaceIndexerAddInApi.cs

API object definition.

{% code-snippet { file-name: face-indexer/FaceIndexerAddInApi.cs } %}

This add-in exposes the API for 3rd parties. *IndexFaces* method is an out-of-process API call and can be used with the following snippet:

~~~ cs
var count = addIn.IndexFaces(assm);
Console.WriteLine($"Indexed {count} face(s)");
~~~

As the result the performance dropped in almost hundred times:

{% include img.html src="stand-alone-result.png" width=300 alt="Result from calling the add-in API from stand-alone application" align="center" %}

Using [ISldWorks::CommandInProgress](http://help.solidworks.com/2016/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks~CommandInProgress.html) SOLIDWORKS API property improves things a little bit, but still performance is dropped more than 10 times compared to the baseline results.

~~~ cs
app.CommandInProgress = true;
var count = addIn.IndexFaces(assm);
app.CommandInProgress = false;
Console.WriteLine($"Indexed {count} face(s)");
~~~

Below is a comparison table of results. Results may vary depending on the size of the assembly and API calls being used.

| Environment                     | Result, seconds | Ratio, % |
|---------------------------------|-----------------|----------|
| Add-In In-Process               | 2.63            | 1        |
| Stand-Alone                     | 241.95          | 92       |
| Stand-Alone Command In Progress | 36.14           | 13.74    |
| VBA Macro                       | 2.57            | 0.98     |
| VBA Macro In-Process Invoking   | 2.20            | 0.84     |
| Stand-Alone In-Process Invoking | 1.77            | 0.67     |

The best performance is gained when add-in API is invoked as in-process call from stand-alone application. This functionality can be achieved by providing deferred call to index faces. This call would put the request into the queue and return the control immediately. The request then will be processed in add-in. [OnIdle](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dsldworksevents_onidlenotifyeventhandler.html) SOLIDWORKS API notification can be used to process the queue. As this event is handled in-process, the actual API calls will be handled in process as well.

It is also important to register the callback which can be called by the add-in to notify the stand-alone application that operation is completed.

Below is an example of the stand-alone application invoking the add-in API in-process.

## Stand-Alone Application

C# application calling the add-in function.

### FaceIndexerCallback.cs

Callback function which notifies the stand-alone application when in-process call is completed. This must be registered as COM object.

{% code-snippet { file-name: stand-alone/FaceIndexerCallback.cs } %}

### Program.cs

Console application invoking the in-process call to add-in API and awaiting result in the callback.

{% code-snippet { file-name: stand-alone/Program.cs } %}

It can also be invoked from the macro or any other type of applications.

## VBA Macro

VBA macro to call the add-in API. In this example User Form is used to keep macro running until the callback function is called.

{% include img.html src="macro-solution-tree.png" width=250 alt="Project tree in VBA macro" align="center" %}

### Macro Module

Main module which is starting the user form

{% code-snippet { file-name: Macro.vba } %}

### FaceIndexerCallback Class Module

Implementation of callback class to receive the notification of completion

{% code-snippet { file-name: FaceIndexerCallback.vba } %}

### Form1 Form

User form to connect to add-in and call its API

{% code-snippet { file-name: Form1.vba } %}

Source code can be downloaded from [GitHub](https://github.com/codestackdev/solidworks-api-examples/tree/master/swex/add-in/face-indexer)