---
layout: article
title: Variables, Constants and Data Types in Visual Basic
caption: Variables, Constants and Data Types
description: Explanation of variables and variable types in Visual Basic
image: /images/codestack-snippet.png
order: 2
---
This section explains the usage of variables and constants when developing code in Visual Basic. 

Variables are used to store the information of different types (e.g. numeric, text, date etc.). Value of the variable can be assigned and changed at run-time.

Unlike variables, constants cannot be changed during the runtime and will always have the value assigned to the constant at declaration. Constant is a good way to store the data which will never change within the application.

There are different types of data which variables and constants can hold. The data types can be classified into two main groups: value types (variable holds the actual value in its own memory allocation) and reference types (variable holds the reference to the object which is stored in different memory allocation). The reference is assigned using **Set** keyword.

> Note. Assigning the reference from one variable to another doesn't copy the value of the object and those variables will point to the same object in memory.