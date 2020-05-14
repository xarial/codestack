---
layout: sw-macro-fix
title: Fix SOLIDWORKS macro which is using future version APIs
caption: Macro Is Using Future Version APIs
description: Fixing the macro which fails when run on old (not the latest) version of SOLIDWORKS and Run-time error '438' - object doesn't support this property or method or Run-time error '445' - object doesn't support this action error is displayed
image: /solidworks-api/troubleshooting/macros/future-version-apis/object-doesnt-support-this-action.png
labels: [macro, troubleshooting]
redirect-from:
  - /2018/04/macro-troubleshooting-macro-using-future-version-apis.html
---
## Symptoms

Recently developed SOLIDWORKS macro is run on old (not the latest) version of SOLIDWORKS. When run, *Run-time error '438': object doesn't support this property or method* is displayed.

![Run-time error '438': object doesn't support this property or method displayed when running the macro](object-doesnt-support-this-property-or-method.png){ width=400 height=151 }

Alternatively *Run-time error '445': object doesn't support this action* can be displayed.

![Run-time error '445': object doesn't support this action is displayed when running the macro](object-doesnt-support-this-action.png){ width=400 height=171 }

## Cause

SOLIDWORKS is [backward compatible](https://en.wikipedia.org/wiki/Backward_compatibility) system which means that older versions of the files or APIs will be compatible with every new release. However SOLIDWORKS is not [forward compatible](https://en.wikipedia.org/wiki/Forward_compatibility) which means that new APIs cannot be used in the older versions of the software. Every release SOLIDWORKS is adding new APIs to the libraries which can be used by the developer to write macros. But those macros cannot be used in the older versions of SOLIDWORKS

## Resolution

* Check SOLIDWORKS API help for the method accessibility which is highlighted by the error

![Availability option in SOLIDWORKS API Help Documentation](comp-config-properties-availability.png){ width=400 height=216 }

* If the earliest available version is newer than it is required to replace the method with an alternative one

Usually SOLIDWORKS names the method with an index, e.g. OpenDoc4, OpenDoc5, OpenDoc6 which indicates the superseded version. If this is the case try to see if there is an older version of this method available. If so this can be used. Please note that older version might have different sets of parameters so it is not always enough just to change the version number

![Difference between versions of the CompConfigProperties API method](comp-config-prps-vers-diff.png){ width=400 height=122 }

* If no older methods available it will be required to overwrite the logic of the macro using alternative methods.
* Upgrade SOLIDWORKS software to the never minimum supported version

Example macro which is using the API added to SOLIDWORKS 2017

{% code-snippet { file-name: suppress-component-sw2017.vba } %}

Modified macro which enables compatibility with SOLIDWORKS 2005 onwards

{% code-snippet { file-name: suppress-component-sw2005.vba } %}
