---
layout: article
title: Tracking objects by temp and persistent ids in SOLIDWORKS API
caption: Tracking Objects
description: This collection of articles explaining how to track different objects while geometry manipulation or across sessions
order: 13
labels: [track, id, persist, reference]
---
When developing application which interacts with SOLIDWORKS entities in some cases it is required to reference certain objects and track them during different actions. For example it is required to find the specific feature in the template model or identify the user selected face after the face got modified (split or merged).

There multiple different ways described below which provides the functionality to tag and track different elements using SOLIDWORKS API.

### Persistent Reference Ids

Allows to retrieve the persistent id for any selectable object in SOLIDWORKS model. The element can be quickly looked up via [IModelDocExtension::GetObjectByPersistReference3](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldocextension~getobjectbypersistreference3.html) SOLIDWORKS API method to get the pointer by the id. The main cons of this method is a size of the id which varies around 250 bytes per entity. So if it is required to track thousands of elements this might not be the ideal approach due to the memory consumption.

Refer [Persistent Reference Id](solidworks-api/document/tracking-objects/persist-references/) article for more information

### Internal Ids

Ids for various group of elements (sketch elements, features, etc.). Internal id only consumes small amount of memory (represented as 1 or 2 Integer or Long values). It is however not possible to lookup the element by its internal id, so it is not suitable for the software where it is required to have an instance access to the object by its id.

Refer [Internal IDs](solidworks-api/document/tracking-objects/internal-ids/)  article for more information

### Tracking Ids

Assignable by the API and used to track the entities (faces, edges and vertices) across modelling operations. For example user selects face on the input body, this body is copied and changed (e.g. split or merged). In this case tracking id will be maintained and all split entities will inherit the id of the parent face.

Refer [Tracking IDs](solidworks-api/document/tracking-objects/tracking-ids/)  article for more information

### Names

Names are available for the user to view and edit via GUI. As names can be easily changed they shouldn't be used as a reliable way of tracking the entity. The names are good for use in the software which is using/modifying template models.

Refer [Object Names](solidworks-api/document/tracking-objects/names/)  article for more information

### Attributes

Attributes are specific features which can be created by API and added to the feature tree. Optionally attribute can be associated with selectable object which allows tracking. Unlike macro features, attributes are native features and will remain functional in the environments where the application which created attributes is not installed.

Refer [Attributes](solidworks-api/data-storage/attributes/) article for more information

Refer the comparison table below which categorizes all approaches above by the following criteria:

* *Lifetime* - how long the id is available
* *Size* - memory consumed by id for a single element
* *Visible* - is this id visible to the user
* *Changeable* - can the id by changed by the user or API
* *Searchable* - can the reference be retrieved directly from id without the need of traversing of all elements
* *Auto Disposable* - is the id disposed automatically when the parent element is destroyed (e.g. deleted)

|Tracking Type|Lifetime|Size|Visible|Changeable|Searchable|Auto Disposable|
|---|---|---|---|---|---|---|
|Persistent Reference Ids|Persistent|~250 bytes|No|No|Yes|Yes|
|Internal Ids|Persistent|2-8 bytes|No|No|No|Yes|
|Tracking Ids|Temp until rebuild|2 bytes|No|No|No|Yes|
|Names|Persistent|usually 10-20 bytes|Yes|Yes|Yes|Yes|
|Attributes|Persistent unless deleted|~1 kilobyte|Can be hidden or visible|Yes|Yes|No|