---
layout: issue-fix
title: Fixing the error when running legacy SWBasic (*.swb) SOLIDWORKS macro
caption: SWBasic (*.swb) macro error
description: Fixing the error when running or editing the legacy macro in swb format
lang: en
image: /solidworks-api/troubleshooting/macros/swb-macro-error/swbasic-swb-macro-filter.png
labels: [macro, troubleshooting]
---
## Symptoms

{% include img.html src="swbasic-swb-macro-filter.png" alt="Selecting the SWBasic Macros (*.swb)" align="center" %}

Legacy SOLIDWORKS macro in *.swb format failed when edited with a 'Compile Error: User-defined type not defined' error. It usually runs correctly if run from the Tools->Macro->Run menu:

{% include img.html src="swb-macro-user-defined-type-not-defined-error.png" width=300 alt="Compile Error: User-defined type not defined" align="center" %}

## Cause

SWBasic macros are scripts stored in ASCII format (i.e. plain text) which cannot store any references information. SOLIDWORKS types are defined in the SOLIDWORKS type libraries which ae not referenced by default in SWBasic macros.

## Resolution

* Open the macro for editing (Tools->Macro->Edit)
* Navigate to the *Tools->References* menu

{% include img.html src="vba-tools-references.png" width=300 alt="References menu in VBA editor" align="center" %}

* Check all SOLIDWORKS type libraries

{% include img.html src="vba-sw-references.png" width=300 alt="SOLIDWORKS Type Libraries in the VBA References dialog" align="center" %}

* Save the macro in *.swp format
