---
layout: article
title: Handling the long operation progress using progress bar in SOLIDWORKS API
caption: User Progress Bar
description: Displaying the long operation progress using user progress bar in SOLIDWORKS API
lang: en
image: /solidworks-api/application/frame/user-progress-bar/taskbar-progress.png
labels: [progress,user progress bar,background]
---
To improve the user experience of your macro or add-in it is recommended to display and update the progress bar when the long SOLIDWORKS API operation is performed.

SOLIDWORKS API provides a built-in method to display the progress while main thread is locked (i.e. operations are performed in process). Progress value and message can be handled via [IUserProgressBar](https://help.solidworks.com/2017/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IUserProgressBar.html) SOLIDWORKS API interface.

Message and progress is displayed in the standard SOLIDWORKS progress bar in the bottom left corner of the application.

{% include img.html src="user-progress-bar.png" alt="Progress and message displayed in the progress bar" align="center" %}

Progress is also reflected in the SOLIDWORKS icon in the task bar.

{% include img.html src="taskbar-progress.png" alt="Progress is displayed in the SOLIDWORKS icon in the task bar" align="center" %}

## Notes and limitations

* Progress values and messages can be overridden by standard progress messages from SOLIDWORKS (e.g. rebuild operation, file load etc.)

## Running the macro

* Open part document with bodies
* Macro traverses all faces of the body and performs data extraction of each face
* Operation is repeated as specified in *ITERATIONS_COUNT* constant
* Progress bar is displayed
* Press ESC to have an option to cancel the operation

{% include_relative Macro.vba.codesnippet %}
