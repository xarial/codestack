---
layout: article
title: Registering add-in using SOLIDWORKS PDM Administration takes long time
caption: Registering the add-in is too slow
description: Troubleshooting the performance issue while registering add-in in SOLIDWORKS PDM administration utility.
image: /solidworks-pdm-api/troubleshooting/addins/slow-addin-registering/server-busy.png
labels: [pdm add-in, error]
---
## Symptoms

It takes very long time to register the add-in in SOLIDWORKS PDM vault using the SOLIDWORKS PDM administration utility. Sometimes the *This action cannot be completed because the '' program is not responding. Choose "Switch To" and correct the problem* message is displayed multiple times while adding the add-in.

{% include img.html src="server-busy.png" alt="Program is not responding error" align="center" %}

In some cases unexpected code is executed or random errors appear.

## Cause

Too many COM visible classes present in the add-in dll.

When add-in is added to the vault via SOLIDWORKS PDM Administration tool, all public COM visible classes from all dlls will be probed by SOLIDWORKS PDM. It means that instances of all classes will be created regardless the class is add-in or not or is it used or not.

In most cases this issue is caused if **Make assembly COM-visible** option is checked in the project properties.

{% include img.html src="assembly-com-visible.png" width=450 alt="COM Visible assembly" align="center" %}

This will make all the public classes automatically visible for COM.

### Example

* Create new PDM add-in and add new class (e.g. MyComClass) and display the message box from within its constructor

~~~cs
public class MyComClass
{
    public MyComClass()
    {
        MessageBox.Show("MyComClass");
    }
}

public class PdmAddIn : IEdmAddIn5
{
    ...
}
~~~

* Check **Make assembly COM-visible** option in the project settings
* Compile the add-in and add this to the vault using the SOLIDWORKS PDM Administration tool.

As the result it would take more time and the following message will be shown which indicates that an instance of **MyComClass** was created while adding the add-in to the vault.

{% include img.html src="message-box.png" width=450 alt="Message box in the COM visible class displayed while registering the add-in" align="center" %}

## Resolution

Do not use the **Make assembly COM-visible** option unless explicitly required.

Only mark the main add-in class with [ComVisible](https://docs.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.comvisibleattribute) attribute with value set to *True*

~~~cs
[ComVisible(true)]
public class PdmAddIn : IEdmAddIn5
{
}
~~~
