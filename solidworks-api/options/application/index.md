---
title: Managing system options (application level) using SOLIDWORKS API
caption: Application Options
description: Collection of articles and examples which demonstrate how to control application (system) options (user preferences) using SOLIDWORKS API
labels: [document, preferences, options]
---
System or application level options are settings available in the options dialog of SOLIDWORKS. Those values can be controlled with following SOLIDWORKS API:

For extracting the values of current options:

* [ISldWorks::GetUserPreferenceDoubleValue](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks~GetUserPreferenceDoubleValue.html)

* [ISldWorks::GetUserPreferenceIntegerValue](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks~GetUserPreferenceIntegerValue.html) 

* [ISldWorks::GetUserPreferenceStringValue](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks~GetUserPreferenceStringValue.html)

* [ISldWorks::GetUserPreferenceToggle](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks~GetUserPreferenceToggle.html)

For changing the options values:

* [ISldWorks::SetUserPreferenceDoubleValue](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks~SetUserPreferenceDoubleValue.html)

* [ISldWorks::SetUserPreferenceIntegerValue](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks~SetUserPreferenceIntegerValue.html) 

* [ISldWorks::SetUserPreferenceStringValue](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks~SetUserPreferenceStringValue.html)

* [ISldWorks::SetUserPreferenceToggle](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks~SetUserPreferenceToggle.html)

This section contains macros and code examples for managing (reading, writing, copying) of various application level system options using SOLIDWORKS API.