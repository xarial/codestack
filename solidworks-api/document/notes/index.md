---
layout: article
title: Working with Notes using SOLIDWORKS API
caption: Notes
description: Collection of articles and code examples about automation of SOLIDWORKS notes annotations
order: 9
image: /images/codestack-snippet.png
---
[INote](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.INote.html) interface is a SOLIDWORKS API representation of the note annotation. This interface would work with notes in assembly, part and drawing environments.

Pointer to the note can be retrieved via [IAnnotation::GetSpecificAnnotation](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iannotation~getspecificannotation.html) SOLIDWORKS API call.

This section contains various macro examples and code snippets for managing notes in SOLIDWORKS using API.