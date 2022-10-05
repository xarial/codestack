---
caption: Format Text Of The Note
title: Example to format portions of the note text with different formats
description: VBA example to apply font effects and styles for different portions of the note in SOLIDWORKS document
image: note-text.png
---
This VBA example demonstrates how to insert note into SOLIDWORKS document and format individual lines with different font effects and styles.

![Formatted text of the note](note-text.png)

Portions of the text can be formatted with **\<FONT\>** instruction. This instruction has 2 attributes

* **effect** - can be equal to **U** (underline) or **RU** (remove underline)
* **style** - can be equal to **B** (bold), **RB** (remove bold), **I** (italic) or **RI** (remove italic)

All the text after the **\<FONT\>** instruction will be formatted according to the value of **effect** and **style**. 

[INote::GetText](https://help.solidworks.com/2023/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.INote~GetText.html) methods returns the resolved value of the note. For the note above it will return the following result:

~~~
First Line Underline
Second Line Bold
Third Line Italic
~~~

[INote::PropertyLinkedText](https://help.solidworks.com/2023/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.INote~PropertyLinkedText.html) property sets or gets the text supporting the **\<FONT\>** instruction. For the note above it will return the following result:

~~~
<FONT effect=U>First Line Underline
<FONT style=B effect=RU>Second Line Bold
<FONT style=RB><FONT style=I>Third Line Italic
~~~

{% code-snippet { file-name: Macro.vba } %}