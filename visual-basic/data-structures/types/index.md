---
layout: article
title: User Defined Types in Visual Basic
caption: Types
description: Article explaining usage of custom user defined types (aka Structures) in Visual Basic
image: /visual-basic/data-structures/types/type-definition-intellisense.png
---
{% include img.html src="type-definition-intellisense.png" width=350 alt="User defined type in intelli-sense" align="center" %}

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

{% include img.html src="type-properties-intellisense.png" width=250 alt="Properties of the user defined type displayed in the intelli-sense" align="center" %}

It is not possible to set the access modifiers or add any functions or procedures within the type:

{% include img.html src="statement-invalid-type-block.png" width=350 alt="Compile Error: Statement invalid inside Type block" align="center" %}

{% code-snippet { file-name: declaration-usage.vba } %}
