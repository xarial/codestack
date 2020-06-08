---
title: Read all invisible custom properties using SOLIDWORKS Document Manager API
caption: Read All Invisible Custom Properties
description: VBA example to read and output all invisible custom properties from the specific model using SOLIDWORKS Document Manager API
labels: [invisible, custom property]
---
SOLIDWORKS models contain several invisible custom properties, such as $PRP:"SW-File Name", $PRP:"SW-Title". Those are read-only and cannot be modified.

This VBA macro reads and outputs all invisible custom properties from the specified model using SOLIDWORKS Document Manager API. The result is output to the immediate window of the VBA editor in the following format:

~~~
...
SW-Short Date: 12/09/2019 [Text]
SW-Long Date: Thursday, 12 September 2019 [Text]
SW-Configuration Name: A [Text]
...
SW-Created Date: Tuesday, 10 September 2019 10:46:55 AM [Text]
SW-Last Saved Date: Thursday, 12 September 2019 8:33:04 PM [Text]
SW-Last Saved By: artem.taturevych [Text]
...
MyProperty: MyValue [Text]
~~~

Specify the file to read properties from in *FILE_PATH* constant.

{% code-snippet { file-name: Macro.vba } %}