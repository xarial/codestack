---
title: User Defined Types in Visual Basic
caption: Types
description: Article explaining usage of custom user defined types (aka Structures) in Visual Basic
image: type-definition-intellisense.png
---
![User defined type in intelli-sense](type-definition-intellisense.png){ width=350 }

In Visual Basic complex data structure (group) of variables can be defined using the **Type - End Type** code block.

~~~ vb
Type MyType
    Var1 As Double
    Var2 As String
End Type
~~~

This enables developers to create easy to understand and use data structures.

Variables of any type can be defined inside the type code block.

Properties declared in type are public and browsable within the intelli-sense:

![Properties of the user defined type displayed in the intelli-sense](type-properties-intellisense.png){ width=250 }

It is not possible to set the access modifiers or add any functions or procedures within the type:

![Compile Error: Statement invalid inside Type block](statement-invalid-type-block.png){ width=350 }

{% code-snippet { file-name: declaration-usage.vba } %}
