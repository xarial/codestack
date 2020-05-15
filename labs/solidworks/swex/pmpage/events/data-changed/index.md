---
layout: article
title: SOLIDWORKS Property Manager Page data changed events handling
caption: Data Change
description: Overview of events associated with data change of SOLIDWORKS property manager page handled in SwEx.PMPage framework
toc-group-name: labs-solidworks-swex
order: 2
---
SwEx framework provides event handlers for the data changes in the controls. Use this handlers to update preview or any other state which depends on the values in the controls.

### Post data changed event

[PropertyManagerPageHandlerEx::DataChanged](https://docs.codestack.net/swex/pmpage/html/E_CodeStack_SwEx_PMPage_PropertyManagerPageHandlerEx_DataChanged.htm) event is raised after the user changed the value in the control which has updated the data model. Refer the bound data model for new values.

{% code-snippet { file-name: Events.DataChanged.* } %}