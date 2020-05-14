---
layout: article
title: Fix 'Please select at least one DLL implementing the IEdmAddIn5 interface' error
caption: Please select at least one DLL implementing the IEdmAddIn5 interface
description: Troubleshooting the 'Please select at least one DLL implementing the IEdmAddIn5 interface' error when registering SOLIDWORKS PDM add-in
image: /solidworks-pdm-api/troubleshooting/addins/no-addin-interface/no-addin-dll.png
labels: [pdm add-in, error]
---
## Symptoms

The following error is shown when adding the add-in with SOLIDWORKS PDM administration tool: *Please select at least one DLL implementing the IEdmAddIn5 interface*

{% include img.html src="no-addin-dll.png" width=450 alt="Error when adding the add-in" align="center" %}

## Cause

Error happens when SOLIDWORKS PDM cannot find the class which implements the [IEdmAddIn5](https://help.solidworks.com/2019/English/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmAddIn5.html) which corresponds to the add-in.

In order for the add-in class to be visible to SOLIDWORKS PDM, it must be public and com visible.

Examples of incorrect declaration of add-in

### Class is not marked as COM Visible

~~~cs
public class PdmAddIn : IEdmAddIn5
{
}
~~~

### Class doesn't have access modifiers (private by default)

~~~cs
[ComVisible(true)]
class PdmAddIn : IEdmAddIn5
{
}
~~~

### Class marked as internal

~~~cs
[ComVisible(true)]
internal class PdmAddIn : IEdmAddIn5
{
}
~~~

## Resolution

Make sure that add-in class is public and decorated with [ComVisible](https://docs.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.comvisibleattribute) attribute with value set to *True*

~~~cs
[ComVisible(true)]
public class PdmAddIn : IEdmAddIn5
{
}
~~~

