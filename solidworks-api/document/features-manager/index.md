---
layout: article
title: Features Manager Tree automation using SOLIDWORKS API
caption: Features Manager
description: Collection of guides for automating features creation via SOLIDWORKS API
lang: en
order: 4
image: /solidworks-api/document/features-manager/feature-manager-api.png
---
{% include img.html src="feature-manager-api.png" alt="Automating features creation via API" align="center" %}

SOLIDWORKS API enables creation of features and automation of Feature Manager Tree via [IFeatureManager](http://help.solidworks.com/2013/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeatureManager.html) interface which can be accessed via [IModelDoc2::FeatureManager](http://help.solidworks.com/2013/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldoc2~featuremanager.html) property.

Each individual feature can be created with different method. Please refer the SOLIDWORKS API help documentation for the list of methods. Alternatively you can record the macro while creating the feature to capture the required method.

It is also possible to extend the scope of standard SOLIDWORKS features by implementing custom [macro feature](http://help.solidworks.com/2013/english/api/sldworksapiprogguide/macro_features/overview_of_macro_features.htm). This will have the same look and feel as any standard feature and will allow to

* Modify or add body
* Add dependency features and regenerate the geometry as needed
* Add dimensions
* Store custom parameters

Feature is presented by [IFeature](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeature.html) SOLIDWORKS API interface. Feature has 2 extension objects:

* Specific feature accessible via [IFeature::GetSpecificFeature2](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~GetSpecificFeature2.html) represents the collection of specific methods and properties for this feature (for example [ISketch](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isketch_members.html) is a specific feature for 2D and 3D sketches).
* Feature definition accessible via [IFeature::GetDefinition](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeature~getdefinition.html) represents the feature parameters (i.e. the ones controlled by the user via Property Manager Page). Modified feature parameters must be applied via [IFeature::ModifyDefinition](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeature~modifydefinition.html) method.

Refer [Identify Feature](identify-feature) example for a helper method allowing to find the interface for feature definition and specific type.