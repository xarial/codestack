---
layout: article
title: SOLIDWORKS PDM API example for handling the data card button click
caption: DataCard Button Click
description: Collection of examples and articles explaining how to handle the button click on data card using SOLIDWORKS PDM Professional API
image: /images/codestack-snippet.png
labels: [hooks, button click, datacard]
---
Data cards functionality can be extended using SOLIDWORKS PDM API by providing the custom logic in the button click handler. Similar to other events button click can be handled within the [IEdmAddIn5::OnCmd](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmaddin5~oncmd.html) overload.

When setting up data card user required to assign the special tag in the options which can be then read from the add-in as a comment which allows to identify the specific button.

This section contains code examples for using SOLIDWORKS PDM API and implementing custom behavior on the data card button click.