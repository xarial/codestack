---
layout: sw-tool
title: SOLIDWORKS macro to rename configurations based on custom property
caption: Rename Configurations Based On Custom Property
description: Macro renames all configurations of assembly or part into the value of the specified configuration specific custom property
image: /solidworks-api/data-storage/custom-properties/rename-configurations-based-custom-property/sw-configuration-name.png
labels: [configuration, custom property, rename, solidworks api, utility]
categories: sw-tools
group: Custom Properties
redirect_from:
  - /2018/04/solidworks-api-model-rename-configurations-based-on-custom-prp.html
---
This macro renames all configurations of assembly or part into the value of the specified configuration specific custom property using SOLIDWORKS API.

{% include img.html src="sw-configuration-name.png" height=200 alt="Configuration name in the configuration properties manager page" align="center" %}

* Run the macro and enter the name of the custom property to read the value from
* Macro will traverse all configurations and rename them based on the corresponding value of the configuration specific custom property
* If property doesn't exist in configuration or value is empty - configuration is not renamed  

{% include_relative Macro.vba.codesnippet %}
