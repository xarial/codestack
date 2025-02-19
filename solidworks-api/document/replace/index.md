---
caption: Replace Model
title: VBA macro to replace and reload active SOLIDWORKS document
description: VBA macro to replace active SOLIDWORKS document with new file and reload without closing the file
image: reload-model.svg
---

This VBA macro replaces the file of the active SOLIDWORKS document and reloads it in the current session without closing the file.

This macro can be useful in demo purposes where it is required to show the model changes instantly

Macro composes the replacement file name based on the value of **REPLACEMENT_SUFFIX**, **REPLACEMENT_PREFIX** and **REPLACEMENT_FILE_NAME** constants

If **REPLACEMENT_FILE_NAME** constant is specified the **REPLACEMENT_SUFFIX**, **REPLACEMENT_PREFIX** are ignored. This can be either a relative file name or an absolute path to the replacement file.

If **REPLACEMENT_FILE_NAME** is empty replacement file name is composed by concatenating the valules of  **REPLACEMENT_SUFFIX** and **REPLACEMENT_PREFIX** with the source file name

This macro supports [Macro+](https://cadplus.xarial.com/macro-plus/) and the **REPLACEMENT_FILE_NAME** can be passed as the first argument.

~~~ vb
Const REPLACEMENT_SUFFIX As String = "_replacement" 'suffix of the replacement file
Const REPLACEMENT_PREFIX As String = "_" 'prefix of the replacement file
Const REPLACEMENT_FILE_NAME As String = "replacement_Part1.sldprt" 'file name of the replacement file
~~~

{% code-snippet { file-name: Macro.vba } %}