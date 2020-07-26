---
title: Managing System Options And User Preferences using SOLIDWORKS API
caption: Options
description: Articles and example to control user preferences and system options at document and application levels using SOLIDWORKS API
image: solidworks-options.svg
labels: [preferences, options]
order: 8
---
![SOLIDWORKS options API automation](solidworks-options.svg){ width=250 }

Most of options in SOLIDWORKS are controlled via SetUserPreference SOLIDWORKS API method. There versions of methods to control text, integer, toggle, numeric properties, etc. properties.

The easies way to find the required identifier of the user preference would be using the following [System Options Table](https://help.solidworks.com/2018/english/api/sldworksapiprogguide/overview/system_options_and_document_properties.htm) or by simply [recording the macro](/solidworks-api/getting-started/macros/recording) and changing the target property. The corresponding SOLIDWORKS API call will be written to the macro.

All system options can be grouped by two categories: application level and document level.

This section contains the code examples and macros for automation user preferences and application settings.
