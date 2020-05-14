---
layout: article
title: Enumerations in Visual Basic (VBA)
caption: Enumerations
description: Introduction to enumeration data types (collection of predefined long constants) in Visual Basic
image: /visual-basic/data-structures/enumerators/enum-icon-intellisense.png
---
{% include img.html src="enum-icon-intellisense.png" width=350 alt="Enumerator type in intelli-sense" align="center" %}

Enumerator is a grouped structure of named constants of type [Long]({{ "visual-basic/variables/standard-types#long" | relative_url }})

The main benefit of enumerator vs constant is an ability to group the constant under single data type and allow an automatic incrementing of values.

Enumerators are usually used to declare different options or actions (e.g. add, remove, delete, move, copy etc. ).

### Declaration and assignment of enumerators

Enumerator can be declared using **Enum - End Enum** code block where each constant declared on new line

~~~ vb
Enum SampleEnum_e
    Val1
    Val2
    Val3
End Enum
~~~

Values of constant can be assigned explicitly or implicitly (automatically). First automatic value is 0 and it is incremented by 1 for every next item.

Enumerator is a value type and can be assigned to the variable. It is possible to use enumerator value directly or via enumerator name

~~~ vb
Dim enumVal As SampleEnum_e
enumVal = SampleEnum_e.Val1 'using enumerator name
enumVal = Val1
~~~

>It is recommended to explicitly use the name of the enumerator. It makes the code more readable and resolves the potential ambiguity if another enumerator or variable has the same name.

{% include_relative declaration.vba.codesnippet %}

### Traversing enumerator values

As enumerators are Long constants it is possible to traverse all the items by knowing the first and last one.

Visual basic allows to declare the special enumerators which are not visible in intelli-sense but still valid values. In order to make the item invisible it is required to use underscore _ symbol at the beginning of the name. For example adding [_First] and [_Last] elements at the beginning and the end of the enumerator would allow defining the boundaries of enumerator values for traversing.

{% include img.html src="enum-invisible-elements.png" width=250 alt="Only visible enumerator values displayed in intelli-sense" align="center" %}

{% include_relative traversing.vba.codesnippet %}

### Flag enumerator (multiple options)

Enumerators can be useful to hold multiple options using bitmasks.

This technique allows combining multiple options within one variable using plus + symbol. it is possible to identify if the specific option was set using **And** bitwise operator.

{% include_relative flag.vba.codesnippet %}
