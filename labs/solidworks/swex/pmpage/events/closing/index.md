---
layout: article
title: SOLIDWORKS Property Manager Page closing events handling
caption: Closing
description: Overview of events associated with closing of SOLIDWORKS property manager page handled in SwEx.PMPage framework
toc_group_name: labs-solidworks-swex
order: 1
---
### Pre Closing event
[PropertyManagerPageHandlerEx::Closing](https://docs.codestack.net/swex/pmpage/html/E_CodeStack_SwEx_PMPage_PropertyManagerPageHandlerEx_Closing.htm) event is raised when property manager page is about to be closed.

Framework passes the reason of close and [closing argument](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_Base_ClosingArg.htm) which allows to cancel property manager page closing and display error to the user as a tooltip.

{% include code-tabs.html src="Events.Closing" %}

This event is raised when Property Manager Page dialog is still visible. There should be no rebuild operations performed within this handler, it includes the direct rebuilds but also any new features or geometry creation or modification (with an exception of temp bodies). Note that some operations such as saving may also be unsupported. In general if certain operation cannot be performed from the user interface while property page is opened it shouldn't be called from the closing event via API as well. Otherwise this could cause instability including crashes. Use [Post closing event](#post-closing-event) event to perform any rebuild operations.

In some cases it is required to perform this operation while property manager page stays open. Usually this happens when page supports pining (swPropertyManagerOptions_PushpinButton flag of [swPropertyManagerPageOptions_e](http://help.solidworks.com/2016/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swPropertyManagerPageOptions_e.html) enumeration in [PageOptionsAttribute](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_Attributes_PageOptionsAttribute.htm)). In this case it is required to set the swPropertyManagerOptions_LockedPage flag of [swPropertyManagerPageOptions_e](http://help.solidworks.com/2016/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swPropertyManagerPageOptions_e.html) enumeration in [PageOptionsAttribute](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_Attributes_PageOptionsAttribute.htm). This would enable the support of rebuild operations and feature creation from within the [PropertyManagerPageHandlerEx::Closing](https://docs.codestack.net/swex/pmpage/html/E_CodeStack_SwEx_PMPage_PropertyManagerPageHandlerEx_Closing.htm) event.

### Post closing event

[PropertyManagerPageHandlerEx::Closed](https://docs.codestack.net/swex/pmpage/html/E_CodeStack_SwEx_PMPage_PropertyManagerPageHandlerEx_Closed.htm) event is raised when property manager page is closed.

Use this handler to perform the required operations.

{% include code-tabs.html src="Events.Closed" %}