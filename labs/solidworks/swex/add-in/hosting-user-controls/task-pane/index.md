---
layout: article
title: Hosting user control in SOLIDWORKS Task Pane using SwEx.AddIn framework
caption: Task Pane
description: Hosting WinForms user control in SOLIDWORKS Task Pane panel using SwEx.AddIn framework
toc_group_name: labs-solidworks-swex
order: 1
---
Any [System.Windows.Forms.UserControl](https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.usercontrol?view=netframework-4.8) can be hosted in the Task Pane by calling the [ISwAddInEx.CreateTaskPane](https://docs.codestack.net/swex/add-in/html/Overload_CodeStack_SwEx_AddIn_Base_ISwAddInEx_CreateTaskPane.htm) method.

~~~ cs
MyControlHost ctrl;
var taskPaneView = CreateTaskPane<MyControlHost>(out ctrl);
~~~

Both COM-visible and not COM-visible controls are supported

~~~ cs
public partial class MyControlHost : UserControl
{
    public IssuesControlHost()
    {
        InitializeComponent();
    }
}
...
[ComVisible(true)]
public partial class MyComVisibleControlHost : UserControl
{
    public IssuesControlHost()
    {
        InitializeComponent();
    }
}
~~~

It is recommended to use COM-visible controls when hosting Windows Presentation Foundation (WCF) control in [System.Windows.Forms.Integration.ElementHost](https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.integration.elementhost?view=netframework-4.8) as keypresses might not be handled properly in com-invisible controls.

## Defining Commands

It is possible to define task pane commands to be added as buttons. It is required to declare the enumeration with commands and provides the commands handler.

~~~ cs
public enum TaskPaneCommands_e
{
    Command1
}

...
TaskPaneControl ctrl;
var taskPaneView = CreateTaskPane<TaskPaneControl, TaskPaneCommands_e>(OnTaskPaneCommandClick, out ctrl);
...

private void OnTaskPaneCommandClick(TaskPaneCommands_e cmd)
{
    switch (cmd)
    {
        case TaskPaneCommands_e.Command1:
            //TODO: handle command
            break;
    }
}
~~~

Commands can be attributed with [TitleAttribute](https://docs.codestack.net/swex/common/html/T_CodeStack_SwEx_Common_Attributes_TitleAttribute.htm) and [IconAttribute](https://docs.codestack.net/swex/common/html/T_CodeStack_SwEx_Common_Attributes_IconAttribute.htm) or [TaskPaneIconAttribute](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Attributes_TaskPaneIconAttribute.htm) for specifying the tooltip and icon respectively.

Standard icon can be set by using the [TaskPaneStandardButtonAttribute](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Attributes_TaskPaneStandardButtonAttribute.htm) attribute where the values defined in [swTaskPaneBitmapsOptions_e](https://help.solidworks.com/2012/english/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swTaskPaneBitmapsOptions_e.html?id=483920098ca24c378c00773c02483619) enumeration

Please see the image below for the diagram of elements of Task Pane.

{% include img.html src="task-pane.png" alt="Task Pane control" align="center" %}

1. WinForms User Control hosted in the Task Pane
1. Task Pane button with the custom icon
1. Task Pane button with default icon
1. Task Pane button with standard swTaskPaneBitmapsOptions_Back icon
1. Task Pane button with standard swTaskPaneBitmapsOptions_Next icon
1. Task Pane button with standard swTaskPaneBitmapsOptions_Ok icon
1. Task Pane button with standard swTaskPaneBitmapsOptions_Help icon
1. Task Pane button with standard swTaskPaneBitmapsOptions_Options icon
1. Task Pane button with standard swTaskPaneBitmapsOptions_Close icon
1. Tooltip for Task Pane button
1. Custom icon for Task Pane Tab
1. Default icon for Task Pane Tab
1. Tooltip for Task Pane Tab