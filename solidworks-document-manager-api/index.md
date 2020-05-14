---
layout: article
title: Introduction to SOLIDWORKS Document Manager API
caption: Document Manager API
description: Collection of examples and articles utilizing SOLIDWORKS Document Manager library
image: /solidworks-document-manager-api/document-manager-api.png
labels: [document manager]
redirect-from:
  - /2018/04/document-manager-api-getting-started.html
---
![SOLIDWORKS Document Manager API](document-manager-api.svg){ width=150 }

This section introduces you to SOLIDWORKS Document Manager - powerful stand-alone library supplied by SOLIDWORKS which exposes API and allows to read and write metadata directly from SOLIDWORKS files stream.

### Benefits

* Free for customers or SOLIDWORKS partners (research, solution or gold) on subscription
* Lightweight - library is about 13 MB in size
* Stand-alone - doesn't require SOLIDWORKS to be installed in order to access the files
* Quick - data is accessed directly from the stream without the need to load the full file into the memory

### Supported Functionality

Document Manager has a limited functionality compared to full SOLIDWORKS API. The following list are the main modules supported by Document Manager library.

* Basic
	* Reading/Writing Custom Properties (including configuration specific and summary information)
    * Reading file relationships (assembly Bill of Materials and drawings)
    * Replacing file relationships (components and drawing view references)
    * Reading components transformations in the assembly
    * Reading tables data in models and drawings
    * Reading [3rd party storage data](http://help.solidworks.com/2015/english/api/sldworksapiprogguide/overview/third-party_data_in_solidworks_files.htm)
    * Getting the information about the configurations, drawing views and their properties
	
* Previews
	* Getting 2D image previews from files and configurations
	
* DimXpert
	* Accessing DimXpert dimensions
    * Accessing DimXpert features
	
* Geometry Streams
	* Getting Parasolid geometry

* XML Streams
	* Reading embedded assembly XML data
    * Reading 3D Content Central data
	
* Tesselation
	* Getting the tessellation (triangulation) data (if stored in the model)

### Application

List of possible applications which could be developed with SOLIDWORKS Document Manager API includes but not limited to the following types of software:

* Product Data Management (PDM) or Product Life cycle Management (PLM) application
	* Bill Of Materials
    * Preview
    * Properties
* 3D Viewers for SOLIDWORKS files
* CAD systems with the requirement of importing the SOLIDWORKS files
* Inspection software with requirements to access DimXpert data

Refer the [Getting Started](getting-started) article for an overview of steps required to start development with Document Manager.
