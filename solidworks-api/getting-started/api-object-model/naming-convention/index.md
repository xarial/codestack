---
layout: article
title: Naming for methods and properties in SOLIDWORKS API
caption: Naming Convention
description: Explanation of the naming convention for methods, properties and interfaces in the SOLIDWORKS API object model (i.e. OpenDoc6 vs OpenDoc5)
lang: en
image: /solidworks-api/getting-started/api-object-model/naming-convention/obsolete-api-interface.png
labels: [obsolete,version,number]
---
SOLIDWORKS API (and SOLIDWORKS) are both backward compatible which means that older versions of APIs are compatible with newer releases of SOLIDWORKS. This means that signatures and behaviors of API methods should not be changed when new version is released. For that purpose SOLIDWORKS introduces the revision system for the methods and interfaces names. Whenever new version of API is available it will be added to the class diagram as **MethodName** *Last Revision + 1*. For example [ISldWorks::OpenDoc6](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~opendoc6.html) is a newer version of [ISldWorks::OpenDoc5](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~opendoc5.html) method. While [IModelDoc2](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2.html) is a newer (and current) version of [IModelDoc](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc.html) interface.

## Obsolete Methods And Interfaces

Although SOLIDWORKS is backward compatible and all the versions of the method should work it is recommended to utilize the newest version compatible with the minimum version of the SOLIDWORKS target program should support.

Main reasons for that are:

* Obsolete methods (or any remarks and descriptions) might not be available in the API Documentation. So it might be required to maintain the previous versions of the API help documentation.

{% include img.html src="obsolete-api-interface.png" width=250 alt="Obsolete IModelDoc API Interface" align="center" %}

* It is not always known what was the reason for adding the replacement method. This might happened due to certain bug (or behavior) present in the older version of the method which might introduce unknown side effects for your program if this method is used.

* It might be problematic to request help from support in case of the issues as the first obvious suggestion would be to upgrade methods to new version as older method can be considered as a *void warranty*.
