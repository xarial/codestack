---
layout: article
title: Usage of IEdmVault5 interface in SOLIDWORKS PDM API
caption: Vault
description: Collection of code examples and tutorials for usage of IEdmVault5 interface in SOLIDWORKS PDM API
order: 2
---
[IEdmVault5](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmvault5.html) interface is a root object in the SOLIDWORKS PDM API object model. It provides an access to the base services of the system, such as:

* User Management
* Batch operations
* Data card management
* Workflow management

In the stand-alone application pointer to the vault can be created via constructor. To initialize the variable it is required to call the [IEdmVault5::Login](http://help.solidworks.com/2018/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmVault5~Login.html) or [IEdmVault5::LoginAuto](http://help.solidworks.com/2018/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmVault5~LoginAuto.html). The first method requires to enter all credentials while second one provides the integrated login, i.e. default SOLIDWORKS PDM login page is displayed or user can be automatically logged in.

Pointer to the initialized instance of [IEdmVault5](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmvault5.html) is passed to the [IEdmAddIn5:OnCmd](http://help.solidworks.com/2018/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmaddin5~oncmd.html) overload when creating the SOLIDWORKS PDM add-in so it is not required to perform the login operation on that object in this case.