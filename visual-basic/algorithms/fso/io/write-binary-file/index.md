---
layout: article
title: Writing the binary content into the file using Visual Basic (VBA)
caption: Write Binary File
description: Function to write the binary content of byte array into a file using Visual Basic (VBA)
image: /images/codestack-snippet.png
labels: [write,text,output]
---
This code snippet demonstrates how to write binary data from the variable of type *Byte()* into the specified file using Visual Basic (VBA).

The below snippet will overwrite the data in the destination binary file

~~~ vb
Dim arr(5237) As Byte
arr(0) = 12: arr(1) = 1: arr(2) = 0
...
WriteByteArrToFile("C:\MyFolder\MyFile.dat")
~~~

Code will automatically create new file if it doesn't exist.

Exception will be thrown in case of any error (for example file cannot be accessed for writing).

{% include_relative WriteBinary.vba.codesnippet %}