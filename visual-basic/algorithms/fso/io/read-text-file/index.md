---
layout: article
title: Reading the content of text file using Visual Basic (VBA)
caption: Read Text File
description: Reading the content of text file into the variable using Visual Basic (VBA)
image: /images/codestack-snippet.png
labels: [read,input]
---
The below code snippet demonstrates how to read the text content from the specified file.

~~~ vb
Dim content As String
content = ReadText("C:\MyFolder\MyFile.txt")
~~~

Code will generate an exception if file doesn't exist or cannot be read.

{% include_relative ReadText.vba.codesnippet %}