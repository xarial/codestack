---
layout: article
title: Variables Scope in Visual Basic
caption: Variables Scope
description: Article explaining different variables scopes (public and private class level, local) in Visual Basic
image: /visual-basic/variables/scope/custom-module-named-module1.png
order: 4
---
In Visual Basic language variables can be declared in different scopes (module level, local in the function, within the code block) with different access modifiers (private or public) which define their visibility in the code.

### Class or module (aka members)

Usually declared in the header of the class (outside of functions and procedures). Member variables have 2 levels of visibility (private and public).

#### Private

Declared using *Dim* keyword and only visible by the function and procedures within the scope of this class or module.

*Module1*
~~~vb
Dim member As Integer
Sub Init()
    member = 10 'visible within the Init function of the Module
End Sub
~~~

Private variables cannot be accessed from outside of the module or class. Compile error: method or data member not found occurs when attempting to access the private member from outside class or module.

{% include img.html src="not-found-member-on-private-variable.png" width=500 alt="Compile error: method or data member not found when calling the privately declared variable from outside class" align="center" %}

#### Public

Public variables declared using the *Public* keyword and can be accessed both from current module or class and external module or class

*Module1*
~~~vb
Dim publicMember As Integer
~~~

*Module2*
~~~vb
Sub main()
Module1.publicMember = 20 'variable is accessed and assigned from the external module
End Sub
~~~

### Local

Local variables declared in the scope of specific code block or function and only visible within that block for the code appearing after the declaration of the variable

~~~ vb
Dim var1 As Double
var1 = 0.5 'var1 is visible here
var2 = 0.25 'var2 is not visible at this step as it is declared in the next line
Dim var2 As Double
~~~

#### Code block
Variables defined in the loops or conditional statements

~~~ vb
If res Then
    Dim localVar As Integer 'local variable defined within the If code block
    localVar = 25
End If
~~~

Although, local code block variables are only visible within this code block, same variable name cannot be used in other code block.

#### Function or procedure

Variables defined in the context of function or procedure. These variables are visible within the function or any nested code blocks.

~~~ vb
Sub main()
    Dim localVar As String
    Dim localBoolVar As Boolean
    localVar = "Hello World"
    If localBoolVar Then 'local variable accessed in the conditional statement
        localVar = "New Hello World" 'local variable modified within the body of conditional statement
    End If
End Sub
~~~

#### Function or procedure parameter

Variables are defined in the signature of the function. These variables declaration equivalent to the local function variable declared at the first line of the function body, i.e. the variables are visible for all code blocks within this function.

~~~ vb
Sub DoWork(paramVar As Double, paramVar2 As Integer)
End Sub
~~~

This example demonstrates the behaviour of variables declared in the different scopes:

*Module1*
{% include img.html src="custom-module-named-module1.png" alt="Module1 module in the Visual Basic Project" align="center" %}

{% code-snippet { file-name: different-scopes-example-module1.vba } %}

*Main Module*

{% code-snippet { file-name: different-scopes-example.vba } %}

Output to [Immediate Window](visual-basic/vba/vba-editor/windows#immediate-window)

{% include img.html src="immediate-window-output.png" width=350 alt="Output of variable values" align="center" %}
