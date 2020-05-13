---
layout: article
title: Reading the content of binary file using Visual Basic (VBA)
caption: Read Binary File
description: Reading the content of binary file into the byte array using Visual Basic (VBA)
lang: en
image: /images/codestack-snippet.png
labels: [read,input,binary]
---
The below code snippet demonstrates how to read the binary content into the variable of type *Byte()* from the specified file in Visual Basic 6 (VBA).

~~~ vb
Dim content As Byte()
content = ReadByteArrFromFile("C:\MyFolder\MyFile.dat")
~~~

Code will generate an exception if file doesn't exist or cannot be read.

{% include_relative ReadBinary.vba.codesnippet %}