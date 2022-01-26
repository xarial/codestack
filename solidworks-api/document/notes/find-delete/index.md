---
layout: sw-tool
title: Macro to find and delete specific notes in the SOLIDWORKS drawing
caption: Find And Delete Notes
description: VBA macro to find and delete notes in all SOLIDWORKS drawing sheets based on the text, regular expressions or empty values
image: delete-note.svg
labels: [note, delete, regular expression, regex]
group: Drawing
---
This VBA macro allows to find and delete all notes in the SOLIDWORKS drawing based on the various criteria, such as by text, expression (property linked text), regular expression or empty values.

## Configuration

Macro can be configured by modifying the constants

~~~ vb
Const FILTER As String = "" 'filter to use whe SEARCH_TYPE is set to ByText or ByExpression
Const SEARCH_TYPE As Integer = SearchType_e.EmptyText 'Type of Search (ByText, ByExpression, EmptyText, EmptyExpression, All)
Const USE_REGULAR_EXPRESSION As Boolean = False 'True to treat value in the FILTER constant as regular expressions
~~~

### Finding All Notes

Set the value of **SEARCH_TYPE** constant to **All** and all notes will be found and deleted

### Searching By Text

Set the value of the display text of the note to the **FILTER** constant and **SEARCH_TYPE** to **ByText** and all notes which match this value will be found and deleted.

### Searching By Expression

Set the value of the expression (property linked text) of the note to the **FILTER** constant and **SEARCH_TYPE** to **ByExpression** and all notes which match this value will be found and deleted.

This can be used to find the notes linked to custom properties, for example the below example will find all notes which are linked to the **Part Number** custom property of the drawing.

~~~ vb
Const FILTER As String = "$PRPSHEET:""Part Number"""
Const SEARCH_TYPE As Integer = SearchType_e.ByExpression
Const USE_REGULAR_EXPRESSION As Boolean = False
~~~

### Searching By Empty Text Or Expression

Set the value of **SEARCH_TYPE** constant to **EmptyText** or **EmptyExpression** and all empty notes will be found and deleted

### Regular Expressions

For more advanced searching options it is possible to use the regular expressions. To enable this option set the **USE_REGULAR_EXPRESSION** to **True**. See [Regular Expressions](https://docs.microsoft.com/en-us/dotnet/standard/base-types/the-regular-expression-object-model) for more information

Example below will find and delete all notes which contain numeric value.

~~~ vb
Const FILTER As String = "\d+"
Const SEARCH_TYPE As Integer = SearchType_e.ByText
Const USE_REGULAR_EXPRESSION As Boolean = True
~~~

{% code-snippet { file-name: Macro.vba } %}
