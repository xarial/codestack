---
title: Determine the type Of active document using SOLIDWORKS API
caption: Determine The Type Of Active Document
description: Example displays the message box of the type of the document currently active in SOLIDWORKS
labels: [assembly, document, drawing, example, part, type]
redirect-from:
  - /2018/03/determine-type-of-active-document.html
---
This example displays the message box of the type of the document currently active in SOLIDWORKS. This will work regardless the document is saved or not.  [IModelDoc2::GetType](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~GetType.html) SOLIDWORKS API method can be used to return the type enumeration which will identify the document as SOLIDWORKS Part, Assembly or Drawing.

{% code-snippet { file-name: Macro.vba } %}