---
layout: article
title: Fix missing Visual Studio SOLIDWORKS API SDK project templates
caption: Project templates are not shown in Visual Studio after the installation of SDK
description: Fixing the installation of SOLIDWORKS API SDK when add-in templates are not shown in Visual Studio C# and VB.NET projects
image: /solidworks-api/troubleshooting/addins/sdk-installation/solidworks-api-sdk-installation.png
labels: [add-in, troubleshooting]
---
## Symptoms
{% include img.html src="solidworks-api-sdk-installation.png" width=250 alt="SOLIDWORKS API SDK installation process" align="center" %}

After installation of SOLIDWORKS API SDK no project templates are shown in Visual Studio New Project dialog for C# and VB.NET

## Cause
System error happened while installing the templates and registering the Wizard assembly in the Global Assembly Cache (GAC). This might be caused by permission issues.

## Resolution

* Locate _SOLIDWORKS API SDK.msi_ file. Usually found in _apisdk_ folder of installation DVD or SOLIDWORKS download package.

* The MSI package is a zip storage. Extract (unzip) the file using any Archive tools which support zip format (such as WinRar, WinZip or 7-Zip)

{% include img.html src="extract-solidworks-api-sdk-files.png" width=500 alt="Unzipping .msi package" align="center" %}

* Find _swvbaddin.zip_ for VB.NET template and _swcsharpaddin.zip_ for C# template in the files of the unzipped folder.

* Open Visual Studio and select _Tools->Options..._ menu command. Navigate to _Projects and Solutions_ group.

{% include img.html src="visual-studio-projects-and-solutions-options.png" width=400 alt="Projects and Solutions options in Visual Studio" align="center" %}

* Copy the path of the _User project templates location_ field value and open this folder in the explorer

* Copy the _swvbaddin.zip_ file into the _Visual Basic_ folder and _swcsharpaddin.zip_ into the _Visual C#_ folder. It is not needed to unzip these files.

{% include img.html src="project-templates-folder.png" alt="Project templates folder" align="center" %}

* Start Visual Studio and create new project. Now C# and VB.NET project templates are available for selection.

{% include img.html src="vbnet-addin-template.png" width=500 alt="VB.NET project template for SOLIDWORKS add-in" align="center" %}

If the following error is displayed when creating the project:

{% include img.html src="solidworkstoolswizard-component-load-error.png" alt="SolidWorksToolsWizard component load error" align="center" %}

* Navigate to the folder unzipped from the _SOLIDWORKS API SDK.msi_ file and find the _SolidWorksToolsWizard.dll_

{% include img.html src="solidworkstoolswizard-dll.png" alt="SolidworksToolsWizard.dll" align="center" %}

* This dll needs to be registered in the Global Assembly Cache (GAC). Navigate to _%windir%\assembly%_ folder in explorer and drag-n-drop the dll file into the folder

{% include img.html src="solidworkstoolswizard-gac.png" alt="SolidWorksToolsWizard dll registered in GAC" align="center" %}

Alternatively follow the [How to: Install an Assembly into the Global Assembly Cache](https://docs.microsoft.com/en-us/dotnet/framework/app-domains/how-to-install-an-assembly-into-the-gac) article on the guide of registering assembly in GAC.
