---
layout: article
title: Dictionary in Visual Basic
caption: Dictionary
description: Explanation of using Dictionary object to store key-value pairs in Visual Basic
image: /visual-basic/data-sets/dictionary/dictionary-key-already-associated.png
order: 3
---
Visual Basic dictionary object is a collection of unique keys and associated values. It is also possible to
[index items with keys in collection]({{ "visual-basic/data-sets/collection#indexing-items-by-keys" | relative_url }}), but in this case it is only possible to have keys of String types. While it is possible to create keys of any type in Dictionary.

Unlike collections dictionaries are COM objects and require reference to *Microsoft Scripting Runtime* library in order to use early binding.

{% include img.html src="microsoft-scripting-runtime-library.png" width=350 alt="Microsoft Scripting Runtime reference" align="center" %}

~~~ vb
Dim dict As Dictionary 'early binding
Set dict = New Dictionary
~~~

It is also possible to use late binding, so it is not required to add the *Microsoft Scripting Runtime* library to the project.

~~~ vb
Dim dict As Object 'late binding
Set dict = CreateObject("Scripting.Dictionary")
~~~

Refer [Early Binding and Late Binding]({{ "visual-basic/variables/declaration#early-binding-and-late-binding" | relative_url }}) article for more information about these approaches.

### Add, edit and traverse elements

In order to add new key-value pair it is required to use **Add** method of **Dictionary** object

~~~ vb
dic.Add <Key>, <Value>
~~~

Keys must be unique otherwise the error will be displayed.

{% include img.html src="dictionary-key-already-associated.png" width=350 alt="Run-time error '457' the key is already associated with an element of this collection when adding the duplicate key" align="center" %}

Elements of the dictionary can be accessed by key or 0-based index either by using () symbol directly on the variable or via **Item** property

~~~ vb
Debug.Print dict.Item(<Key>)
Debug.Print dict(<Key>)
~~~

All keys from the dictionary can be retrieved using the **Keys** property.

All values from the dictionary can be retrieved using the **Values** property.

{% code-snippet { file-name: add-edit-traverse.vba } %}

### Key compare mode

By default the compare mode for keys is set to **Binary** comparison. This means if dictionary has keys of type String the keys are case-sensitive, i.e. it is acceptable to have both *A* and *a* as the key.

**Exists** method provides a safe way to check if the key already registered in the dictionary.

**CompareMode** property allows to set the mode which should be used when comparing the entries.

* BinaryCompare (default). String keys are case-sensitive
* TextCompare. String keys are case-insensitive

Mode can only be changed for an empty dictionary (without values), otherwise the error will be displayed.

{% include img.html src="change-compare-mode-invalid-procedure.png" width=400 alt="Run-time error '5': Invalid procedure call or argument when changing the compare mode of dictionary objects with elements" align="center" %}

{% code-snippet { file-name: exists-compare.vba } %}

### Remove elements

Any element can be removed from the dictionary either by key or by 0-based index using **Remove** method.

>Attempt on removing the item which is not present in the dictionary will throw an exception

{% include img.html src="dictionary-remove-object-error.png" width=250 alt="Run-time error '32811': Method Remove of object 'IDictionary' failed when removing non-existent element" align="center" %}

**RemoveAll** method allows to clear the dictionary and remove all items.

{% code-snippet { file-name: remove.vba } %}
