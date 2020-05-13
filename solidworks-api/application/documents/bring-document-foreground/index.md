---
layout: article
title: Bring document foreground (activate document) using SOLIDWORKS API
caption: Bring Document Foreground (Activate Document)
description: Example demonstrates how to bring the document selected by path to foreground (make active)
lang: en
image: /images/codestack-snippet.png
labels: [activate doc, assembly, example, foreground, open document]
redirect_from:
  - /2018/03/bring-document-foreground-activate.html
---
This example demonstrates how to bring the document selected by path to foreground (make active) using [ISldWorks::ActivateDoc3](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~activatedoc3.html) SOLIDWORKS API.

Document can be opened in 2 states (visible or hidden). Hidden document are usually models which are loaded into the memory from the components in the assembly or drawing. In this case when [ISldWorks::OpenDoc6](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~opendoc6.html) method is called the document will not be brought foreground automatically. Similar scenario applies to closing the document which is loaded as a component: document will be made invisible rather than closed.

* Run the macro when no files are opened - file will be opened and closed
* Open assembly and run the macro. In this case [ISldWorks::OpenDoc6](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~opendoc6.html) API doesn't force the part to be brought foreground, so it is required to force activate it.

[Download sample files](SimpleBox.zip)

{% include_relative Macro.vba.codesnippet %}