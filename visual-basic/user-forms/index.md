---
layout: article
title: User Form And Controls in Visual Basic 6 (VBA)
description: Article explaining the usage of user interface elements (text box, combo box, list etc.) and forms in Visual Basic 6 (VBA)
caption: User Form And Controls
lang: en
image: /visual-basic/user-forms/user-form-shown.png
order: 11
---
User forms allow to defined the custom Graphics User Interface (GUI) to collect user inputs, show outputs or provide an interaction with your application.

User form can be added by calling the *Insert UserForm* command

{% include img.html src="insert-user-form.png" width=350 alt="Insert User Form" align="center" %}

By default forms will be named as *UserForm1*, *UserForm2*, etc., but it is recommended to give forms meaningful names.

## Adding Controls

Form can be customized and additional controls can be placed onto the form.

{% include img.html src="user-form-layout.png" width=450 alt="Layout of the user form" align="center" %}

1. User Form design layout
1. Toolbox with controls
1. Control placed on the form layout
1. Properties of the control

Properties of the controls can be customized.

## Code Behind

Form and its controls are exposing different [events](/visual-basic/events/), such as click, select, mouse move etc.

Event handlers are defined in the code behind of the form.

{% include img.html src="view-code-menu-command.png" height=400 alt="View Code command of User Form" align="center" %}

Available control events can be selected from the drop-down list.

{% include img.html src="windows-form-code-behind.png" width=600 alt="Control events" align="center" %}

~~~vb
Private Sub CommandButton1_Click()
    MsgBox "CommandButton1 Clicked!"
End Sub
~~~

## Displaying Form

Form can be displayed by calling the *Show* method. This method should be called on the variable which equals to the form name. Note, it is not required to declare or instantiate the form variable explicitly (as it is required with the classes). This is done automatically when form is added to the project.

{% include img.html src="user-form-shown.png" alt="User Form" align="center" %}

For can be displayed in 2 modes

### Modal

In this mode form is opened foreground and parent window is not accessible until the form is closed.

~~~vb
Sub main()

    UserForm1.Show

End Sub
~~~

### Modeless

Form is opened in a way that parent window is accessible and not blocked. To open the form in the modeless mode it is required to pass the *vbModeless* parameter to *Show* method.

~~~vb
Sub main()

    UserForm1.Show vbModeless

End Sub
~~~

