---
layout: article
title: Arrays in Visual Basic
caption: Array
description: Article explaining arrays in Visual Basic - set of elements which stored in the single variable and can be accessed by index
image: /visual-basic/data-sets/array/two-dimensional-array.png
order: 1
---
Visual Basic array is a set of elements which is stored in the single variable and can be accessed by index.

In order to declare array it is required to append the variable name with parenthesis () symbol;

~~~ vb
Dim arr() As Double 'declaring the array of doubles
~~~

Array elements can be accessed by the index

~~~ vb
Dim arr(2) As String
Dim elem As String
elem = arr(0) 'getting first element
~~~

### Initializing arrays in Visual Basic

Arrays is a set of the fixed size. Size can be assigned while declaration.

~~~ vb
Dim arr(2) As Double 'declaring the array of 3 doubles (from 0 to 1)
~~~

> The single number between the parenthesis represents the upper boundary of the array. By default arrays are 0-based. So specifying (5) as array size means that there will be 6 elements in the array.

It is possible to explicitly specify the upper and inner boundaries of the array

~~~ vb
Dim arr(1 To 5) As Double 'declaring the array of 5 doubles (from 1 to 5)
~~~

> It is recommended to use 0-based array as it is a common practice across different programming languages.

Array elements can be accessed and changed by their index. If specified index is outside of the boundaries of the array, run-time error is raised.

In some cases size of the array cannot be known at the compile time and it will be identified during the run-time. In this case array can be declared without the size (i.e. not initialized). Such array can be then resized dynamically using the **ReDim** keyword.

~~~ vb
Dim arr() As Double
ReDim arr(2) 'initializing the size of the array
~~~

In order to retrieve the upper and lower boundaries of the array **UBound** and **LBound** can be used respectively.

> UBound function cannot be used to identify if the array is initialized as it will throw an exception when used on not initialized array. Use the **IsArrayInitialized** function from the example below to safely identify the state of the array.

![Run-time error 9: subscript out of range while reading the upper boundary of uninitialized array](subscript-out-of-range.png){ width=350 }

{% code-snippet { file-name: initialize.vba } %}

### Filling array with data

Array elements can be treated as individual variables and the same rules of reading and editing the data applies to array elements as any other variable. Refer [Variables](visual-basic/variables) article for more information.

~~~ vb
Dim arr(2) As Double
arr(<INDEX>) = 10 'changing the value of the variable at <INDEX>
Debug.Pring arr(<INDEX>) 'reading the value of the variable at <INDEX>
~~~

{% code-snippet { file-name: fill.vba } %}

### Resizing array

Array size can be changed at run-time. But that can only be done for the array whose size was not explicitly specified on declaration

~~~ vb
Dim arrCanResize() As Integer 'size of this array can be resized
Dim arrCannotResize(3) As Integer 'size of this array cannot be resized
~~~

Attempt of resizing the already dimensioned array will result in the compile error:

![Compile error: Array already dimensioned](array-already-dimensioned.png){ width=350 }

#### Clearing the existing values

**ReDim** keyword allows to resize the array. In this case all existing values will be cleared.

{% code-snippet { file-name: resize-clear.vba } %}

#### Preserving existing values

In order to keep the existing values of the array it is required to use **ReDim Preserve** keyword.

{% code-snippet { file-name: resize-preserve.vba } %}

### Two dimensional array

Two-dimensional array allows to store the table data and can be declared in the following format

~~~ vb
Dim table(<ROWS COUNT>, <COLUMNS COUNT>) As String
~~~

> 2-dimensional array can be resized, but if it is required to preserve the values only second (column) dimension can be resized.

{% code-snippet { file-name: two-dimensional.vba } %}

![Values of two-dimensional array (matrixArr) in the Watch window from the code above](two-dimensional-array.png)
