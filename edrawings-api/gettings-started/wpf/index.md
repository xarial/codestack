---
layout: article
title: Hosting SOLIDWORKS eDrawings control in Windows Presentation Foundation (WPF)
caption: Hosting Control in WPF
description: Detailed guide on hosting SOLIDWORKS eDrawings control as WPF User Control in Windows Presentation Foundation (WPF)
lang: en
image: /edrawings-api/gettings-started/wpf/edrawings-wpf-window.png
labels: [edrawings,host,wpf]
---
eDrawings API doesn't provide a native WPF control to be used in WPF. It is however possible to use the [WindowsFormsIntegration](https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.integration) framework to host Windows Forms Control in the Windows Presentation Foundation (WPF) environment. Follow [Hosting eDrawings control in Windows Forms]({{ "/edrawings-api/gettings-started/winforms/" | relative_url }}) guide of creating the eDrawings control for Windows Forms.

## Creating new project

* Start Visual Studio
* Create new project and select *WPF Application* in the *Visual C#* templates section
{% include img.html src="visual-studio-new-wpf-project.png" width=550 alt="Creating WPF application" align="center" %}
* Follow the [Hosting eDrawings control in Windows Forms]({{ "/edrawings-api/gettings-started/winforms/" | relative_url }}) guide for steps of adding eDrawings interop
* Add reference to *WindowsFormsIntegration*

## Creating the eDrawings WPF control

Create a wrapper for the eDrawings host Windows Forms control

### eDrawingHost.cs

{% include_relative eDrawingHost.cs.codesnippet %}

Create new WPF User Control which will host eDrawings and can be placed on other WPF controls or WPF windows

The solution tree will be similar to the one below.

{% include img.html src="visual-studio-solution-tree.png" width=350 alt="eDrawings WPF solution tree" align="center" %}

### eDrawingsHostControl.xaml

There will be no logic or additional markup in the XAML of the control and all will be implemented in the code behind

{% include_relative eDrawingsHostControl.xaml.codesnippet %}

### eDrawingsHostControl.xaml.cs

{% include_relative eDrawingsHostControl.xaml.cs.codesnippet %}

In this example the control defines the dependency property *FilePath* which can be bound and represent the path to the SOLIDWORKS file to be opened in the eDrawings

### MainWindow.xaml

Add the following markup to the MainWindow. It defines the text box control whose *Text* property is bound to *FilePath* dependency property of WPF eDrawing control. Which means that the file will be loaded immediately once the value in the text box is changed.

{% include_relative MainWindow.xaml.codesnippet %}

Change the path to file in the text box to see the file loaded into the WPF form.

{% include img.html src="edrawings-wpf-window.png" width=350 alt="SOLIDWORKS file is loaded into the WPF eDrawings control" align="center" %}

Source code is available on [GitHub](https://github.com/codestackdev/solidworks-api-examples/tree/master/edrawings-api/eDrawingsWpfHost)
