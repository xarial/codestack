---
caption: Toggle Drawing Detailing Mode On Save
title: Save SOLIDWORKS drawing with detailing mode on and off
description: VBA Macro to toggle detailing mode on and off while saving
---
When working with large drawings it may be beneficial to employ the detailing mode. In order to properly utilize the detailing mode it is required to save the data within the document itself.

This process may decrease the saving performance.

The toggle option to enable or disable saving of the detailing mode data is driven by document preferences.

This macro allows to turn on or off the settings and perform the saving of the document.

~~~ vb
Const ENABLE As Boolean = True 'True to save with detailing data, False to save without the detailing data
~~~

It is possible to create 2 macro buttons (one which saves with detailing data and one which saves without the detailing data).

{% code-snippet { file-name: Macro.vba } %}