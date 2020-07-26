---
title: Developing add-ins using SOLIDWORKS PDM API
caption: Developing Add-Ins
description: Instructions and code examples for developing add-ins for SOLIDWORKS PDM
labels: [add-in,pdm]
---
Add-ins in SOLIDWORKS PDM are applications which are integrated into the systems. Add-ins are installed into the SOLIDWORKS PDM Administration Console and redistributed among all clients which are connected to the vault.

Add-in enables an access to all available SOLIDWORKS API interfaces and methods.

In order to create an add-in it is required to implement the [IEdmAddIn5](https://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmaddin5.html) interface.

This section provides guidelines of creating and troubleshooting add-ins using SOLIDWORKS PDM API.