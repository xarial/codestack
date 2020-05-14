---
layout: sw-macro-fix
title: How to fix Missing SOLIDWORKS Type Library References error
caption: Missing SOLIDWORKS Type Library References
description: Fixing Can't find project or library error in legacy macro
image: /solidworks-api/troubleshooting/macros/missing-solidworks-type-library-references/error-cant-find-project-or-library.png
labels: [macro, troubleshooting]
redirect-from:
  - /2018/04/macro-troubleshooting-missing-solidworks-type-library-references.html
---
## Symptoms

* Legacy SOLIDWORKS macro downloaded from the internet or have been developed in-house some time ago failed to run.
*Can't find project or library error* is displayed and of the SOLIDWORKS declarations are highlighted.

![Can't find project or library error when running the macro](error-cant-find-project-or-library.png){ width=320 height=182 }

Alternatively non-SOLIDWORKS declarations can be highlighted (such as Left or Mid functions)

![Can't find project or library error on Left function in VBA](error-cant-find-project-or-library-left.png){ width=320 height=185 }

* If the libraries were never selected in the macro the *Compile error: user-defined type not defined* can be displayed.

![Compile error: user-defined type not defined](compile-error-user-defined-type-not-defined.png){ height=200 }

## Cause

* Macro is pointing to older versions of SOLIDWORKS type libraries and cannot resolve them automatically. As the result the libraries are marked as MISSING.

* SOLIDWORKS type libraries were never selected or were explicitly deselected in the macro (this usually happens when macro is converted from *.swp macro)

## Resolution

* Open the macro for [editing](http://help.solidworks.com/2017/english/solidworks/sldworks/t_edit_macro.htm) via Tools->Macro->Edit menu
* Navigate to Tools->References menu in the VBA editor
* Select the SOLIDWORKS type libraries as shown below. If libraries cannot be found in the *Available References* list use *Browse...* button and find the *sldworks.tlb*, *swconst.tlb*, *swcommands.tlb* in the installation folder of SOLIDWORKS.

![Required SOLIDWORKS type libraries](selected-sw-references.png){ height=200 }

* If libraries are selected or **MISSING** keyword is present it is required to force update the references by following the steps below:

![List of missing references in VBA macro](fix-update-vba-references.png){ width=320 height=269 }

* Uncheck all of the libraries which are referencing SOLIDWORKS. (including libraries with **MISSING** keyword next to them)
* Click OK
* Open the same dialog again and check corresponding SOLIDWORKS libraries. Those are usually available in the references list.
If not you can use 'Browse...' button to manually select the libraries from the SOLIDWORKS installation folder

Alternatively you can copy paste all the code to newly created macro.
