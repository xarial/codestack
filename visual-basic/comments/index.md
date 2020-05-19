---
layout: article
title: Comments in Visual Basic
caption: Comments
description: Article explaining the usage of comments to annotate the code
order: 8
---
Comments are free texts which can be placed in the code for annotation and reference purposes. Comments are excluded by the compiler and allow to add any text into it

In Visual Basic, comment is any text placed after the apostrophe **'** symbol and the end of the line. Although the color scheme is adjustable, the default color for comment in Visual Basic is green

Comments can be added at the beginning of the line

~~~ vb
'Function is executing some code
Sub DoWork()
End Sub
~~~

Comments can also be placed at the end of the line

~~~ vb
Dim a As String 'declaring the string variable
a = "Hello World" 'assigning value to string variable
~~~

Although, comments are good tool to annotate your code, try not to overuse the comments as this may make code look very busy. Instead try to use descriptive names for the variable and functions.

> Good code comments itself

Instead of 
~~~ vb
Dim var1 As String
var1 = = "Xarial" 'company name
~~~

use 

~~~ vb
Dim companyName As String
companyName = "Xarial"
~~~

Also avoid adding comments for an obvious code. For example the comment in code below is a tautology and should not be added

~~~ vb
'Calculates the square root of the value
Function CalculateSquareRoot(val as double)
End Function
~~~

In general I would recommend adding the comments in the following cases

* For education and tutorial purposes
* For the code which is not obvious to understand, this might be some complicated algorithm
* For the workaround, i.e. certain functionality can be done in easier way, but there is a known limitation or a bug. For example this can happen where 3rd party APIs are used and there is a known problem with a method
* As a placeholder for future work. In this case you can use a **TODO** placeholder

~~~ vb
Function IsValid() As Boolean
    'TODO: implement validation
    IsValid = True
End Function
~~~