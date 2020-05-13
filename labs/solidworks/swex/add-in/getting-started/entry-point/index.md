---
layout: article
title: Entry Point of SwEx.AddIn framework
caption: Entry Point
description: Instructions on starting the coding with the SwEx.AddIn framework for SOLIDWORKS
lang: en
toc_group_name: labs-solidworks-swex
order: 2
---
## Registering the add-in

In order to register the SOLIDWORKS add-in with SwEx framework it is required:

* Create a public class which inherits the [SwAddInEx](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_SwAddInEx.htm) class
* Make this class com visible by adding the *System.Runtime.InteropServices.ComVisibleAttribute* attribute
* Add the [AutoRegisterAttribute](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Attributes_AutoRegisterAttribute.htm) attribute to add the required information to the registry.

### C\#

{% include_relative EntryPoint.cs.codesnippet %}

### VB.NET

{% include_relative EntryPoint.vb.codesnippet %}

## OnConnect

This function is called within the ConnectToSw entry point. Override this function to initialize the add-in.

Return the result of the initialization. Return *true* to indicate that the initialization is successful. Return 'false' to cancel the loading of the add-in.

This override should be used to validate license (return false if the validation is failed), add command manager, task pane views, initialize events manager, etc.

## OnDisconnect

This function is called within the DisconnectFromSw function. Use the function to release all resources. You do not need to release the com pointers to SOLIDWORKS or command manager as those will be automatically released by SwEx framework.

## Accessing SOLIDWORKS application objects

SwEx framework provides the access to the following objects which are preassigned by the framework

### App property
Pointer to SOLIDWORKS application

### AddInCookie property
Add-in id

### CmdMgr property
Pointer to command manager

## Unregistering add-in
Add-in will be automatically removed and all COM objects unregistered when project is cleaned in Visual Studio