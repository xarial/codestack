---
layout: article
title: Tables (BOM, General, Revision etc.) automation using SOLIDWORKS API
caption: Tables
description: Article explaining the functions to work with tables (Bill of Materials, General, Weldment Cut List, Holes Table) using SOLIDWORKS API
order: 8
---
All table types supported by SOLIDWORKS can be accessed via API. This includes but not limited to

* Bill Of Material
* General
* Weldment Cut List
* Holes Table

etc.

All table inherit the [ITableAnnotation](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ITableAnnotation.html) SOLIDWORKS API interface. This interface provides the method to work with the table (i.e. change cells, change formatting, add/remove rows etc.).

There are specific table annotation for a generic table annotation. For example [IBomTableAnnotation](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IBomTableAnnotation.html) is a specific table annotation for Bill of Materials (BOM) table. Generic table annotation can be cast to specific by directly assigning the pointer.

Table is also present in the Feature Manager tree which means that it also provides methods exposed by the [IFeature](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeature.html) interface. Each specific table annotation provides the property to access the specific table feature. For example [IBomTableAnnotation::BomFeature](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibomtableannotation~bomfeature.html) will return the specific BOM feature. To get the pointer to [IFeature](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeature.html) it is required to call the ::GetFeature method for all specific table features.
