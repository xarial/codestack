---
layout: sw-macro-fix
title: Fixing the error when running legacy SWBasic (*.swb) SOLIDWORKS macro
caption: SWBasic (*.swb) macro error
description: Fixing the error when running or editing the legacy macro in swb format
image: swbasic-swb-macro-filter.png
labels: [macro, troubleshooting]
---
## Symptoms

![Selecting the SWBasic Macros (*.swb)](swbasic-swb-macro-filter.png)

Legacy SOLIDWORKS macro in *.swb format failed when edited with a 'Compile Error: User-defined type not defined' error. It usually runs correctly if run from the Tools->Macro->Run menu:

![Compile Error: User-defined type not defined](swb-macro-user-defined-type-not-defined-error.png){ width=300 }

## Cause

SWBasic macros are scripts stored in ASCII format (i.e. plain text) which cannot store any references information. SOLIDWORKS types are defined in the SOLIDWORKS type libraries which ae not referenced by default in SWBasic macros.

## Resolution

* Open the macro for editing (Tools->Macro->Edit)
* Navigate to the *Tools->References* menu

![References menu in VBA editor](vba-tools-references.png){ width=300 }

* Check all SOLIDWORKS type libraries

![SOLIDWORKS Type Libraries in the VBA References dialog](vba-sw-references.png){ width=300 }

* Save the macro in *.swp format
