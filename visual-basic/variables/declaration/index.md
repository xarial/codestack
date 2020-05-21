---
layout: article
title: Declaration and assignment of variables and constants in Visual Basic
caption: Variables and Constants
description: Overview of declaration of variables and constants and assignment of values in Visual Basic
image: option-explicit-variable-not-defined.png
order: 3
---
## Declaring variables

Variables can be declared either explicitly or implicitly in Visual Basic. To declare variable explicitly it is required to use *Dim* keyword or *Public* keyword to declare the variable as public class or module member (refer the [Variables Scope](/visual-basic/variables/scope) article for more information).

Type of the variable can be declared using *As* keyword.

~~~ vb
Dim textVal As String
~~~

> If type of the variable is not explicitly stated than it is defaulted to [Variant](/visual-basic/variables/standard-types#variant)

~~~ vb
Dim varVal 'implicitly declared as Variant
~~~

*Dim* keyword is not used when variable is declared implicitly. In this case value of the variable assigned directly. 

~~~ vb
implicitVal = 10 'Implicit declaration and assignment of the variable
~~~

This is not recommended approach as it may introduce the ambiguity and potential issues with the code.

> In order to force explicit variable declaration in Visual Basic it is required to use *Option Explicit* statement. In this case compile error will occur if implicit variable is identified.

![Compile error when Explicit option is enabled and implicit variable assignment is used](option-explicit-variable-not-defined.png){ width=300 }

### Declaring group of variables of the same type

Variables can be declared in the group.

~~~ vb
Dim textVar1, textVar2, textVar3 As String '3 variables explicitly declared as String
~~~

This approach allows to make code more readable and compact.

### Declaring group of variables with different types

It is allowed to use declaration characters for each variable to declare the type explicitly using the short declaration

~~~ vb
Dim intVar%, doubleVar#, longVar& 'Integer, Double and Long variables are declared explicitly using short declaration
~~~

Refer the [Standard Types](/visual-basic/variables/standard-types) article for the list of declaration characters.

> This is a legacy way to declare the variables. This approach is not recommended way to declare the variables.

## Assigning the values

In order to assign the value of the variable it is required to use equal (=) sign, where the variable name appears on the left and variable value appears on the right.

~~~ vb
Dim varName As String
varName = "VarValue"
~~~

it is possible to copy the value from one variable to another

~~~ vb
Dim var1 As Integer
Dim var2 As Integer
var1 = 10
var2 = var1 'value of var2 now equals to var1 which equals to 10
~~~

It is possible to assign the value to the variable as the result of calling another function. Refer [Functions and Procedures](/visual-basic/functions) article on more information about functions.

~~~ vb
Dim funcRes As Double
funcRes = GetValueFunc()
~~~

## Declaring constants

Constant allows to define the value which will not change during execution. It is usually used for declaring mathematical constants (e.g. PI, gravitational constant, etc.), conversions factors (e.g. hours to minutes, inches to meters etc.) or any other program specific values.

Constant is declared using *Const* statement and must assign the value on declaration.

~~~ vb
Const G As Double = 9.8 'gravitational constant
~~~

Same as variable declaration constant type can be defined explicitly (using *As* keyword) or implicitly.

Once declared value of the constant cannot be changed. Otherwise the compile error will occur.

![Compile error assignment to constant not permitted when changing the value of constant variable](error-changing-constant.png){ width=300 }

This code example demonstrates different ways of declaring and assigning constants and value variables.

{% code-snippet { file-name: assigning-value-variables.vba } %}

## Assigning reference variables

Unlike value types, [references types](/visual-basic/variables/user-defined-types#class) must follow several additional rules when assigning the value.

![Custom User Defined Type](user-type-declaration.png){ width=200 }

* It is required to use *new* keyword to create new instance of the referenced type. Otherwise Run-time error '91' will be displayed

![Run-time error '91': Object variable or With block variable not set when calling the member of uninitialized reference](error-91-when-calling-member-non-initialized-class.png){ width=350 }

* It is required to use *Set* keyword to assign the value, otherwise the Run-time error '91' will be displayed

![Run-time error '91': Object variable or With block variable not set when Set keyword is not used to assign the reference to the variable](error-when-not-using-set-keyword.png){ width=350 }

See code below for the correct assignment of reference type variable.

{% code-snippet { file-name: assigning-reference-variables.vba } %}

> References variables are only holding the pointer to the actual value, i.e. Set keyword assigns the reference (not the actual value like in value types). That means if reference of one variable is assigned to another variables, both of them now refer the same data.

### Early binding and late binding

Binding is a process of assigning object to a variable. When early binding is used the specific object type is declared in advanced so the binding can occur at compile time. Late binding is resolved at runtime and specific object type is not known in advance.

~~~ vb
Dim objLate As Object 'example of late binding
Dim objEarly as MySpecificType 'example of early binding
~~~

Early bound objects are usually initiated with *new* keyword

~~~ vb
Dim objEarly as MySpecificType
Set objEarly = new MySpecificType
~~~

While late bound objects are usually initiated with [CreateObject](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/createobject-function) or [GetObject](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/getobject-function) functions

~~~ vb
Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
~~~

But it is still acceptable to use *new* keyword in late binding and CreateObject or GetObject in early binding.

Benefits of early binding

* Performance. Compiler can perform required optimization as the type of the object and its size is known at compile time
* Maintainability. Code is cleaner and easier to maintain and read when specific type is declared
* Dynamic help and IntelliSense (code completion) features are available for early bound objects

Benefits of late binding

* No need to maintain 3rd party references which may be an issue when code is ported to another environment or another version of 3rd party references is released. Refer this [Example of references issue](/solidworks-api/troubleshooting/macros/missing-solidworks-type-library-references)
