---
title: 'SOLIDWORKS Macros Troubleshooting: Issues And Resolutions'
caption: 'Macros Troubleshooting: Issues And Resolutions'
description: Overview and solutions for the most common errors of running the macros in SOLIDWORKS
labels: [macro, not working, problem, solidworks api, troubleshooting, vba]
styles:
  - /_assets/styles/catalogue.css
redirect-from:
  - /2018/04/macros-troubleshooting-issues-and-resolutions.html
---
SOLIDWORKS macro is the most common way to automate and extended the functionality of SOLIDWORKS via its API.
Macros can be developed in-house or downloaded from the different web-sites including SOLIDWORKS forum, 3D Content Central, [CodeStack](/solidworks-tools) etc. or even recorded from SOLIDWORKS.

But in some cases the macro doesn't work as expected. This is usually one of the following scenarios:

* Macro which used to work correctly before stopped working.
* Macro works on some of the workstations but not on the others.
* Macro works correctly for some models but not the others.

In this article I will go through the most common symptoms of the errors in the macros.

Browse the errors list to find the most common solutions.

Click link to get the detailed description of the issues, its cause and the steps to resolve the problem.

## Errors List

* Run-time Error '91': Object variable or With block variable not set
  * [Solution 1](/solidworks-api/troubleshooting/macros/assembly-drawing-lightweight-components/)
  * [Solution 2](/solidworks-api/troubleshooting/macros/macro-multiple-entry-points/)
  * [Solution 3](/solidworks-api/troubleshooting/macros/create-sketch-segments-error/)
  * [Solution 4](/solidworks-api/troubleshooting/macros/preconditions-not-met/)
  * [Solution 5](/solidworks-api/troubleshooting/macros/selection-inconsistency/)

* Compile Error: Can't find project or library
  * [Solution 1](/solidworks-api/troubleshooting/macros/missing-solidworks-type-library-references/)

* Run-time error '424': Object required
  * [Solution 1](/solidworks-api/troubleshooting/macros/merged-macro-error/)

* Run-time error '13': Type mismatch
  * [Solution 1](/solidworks-api/troubleshooting/macros/preconditions-not-met/)

* Compile Error: User-defined type not defined
  * [Solution 1](/solidworks-api/troubleshooting/macros/swb-macro-error/)

* Run-time error '438': object doesn't support this property or method
  * [Solution 1](/solidworks-api/troubleshooting/macros/future-version-apis/)

* Run-time error '429': ActiveX component can't create object
  * [Solution 1](/solidworks-api/troubleshooting/macros/missing-com-component/)

* Run-time Error '5': Invalid procedure call or argument
  * [Solution 1](/solidworks-api/troubleshooting/macros/model-title-inconsistency-displaying-extension/)

* Compile error: The code in this project must be updated for use on 64-bit systems is displayed. Please review and update Declare statements and then mark item with the PtrSafe attribute
  * [Solution 1](/solidworks-api/troubleshooting/macros/32-windows-api-functions-incorrect-use/)

* Cannot Open (for VBA macros)
  * [Solution 1](/solidworks-api/troubleshooting/macros/too-long-macro-path/)

* Compile error: Invalid outside procedure error
  * [Solution 1](/solidworks-api/troubleshooting/macros/too-long-vba-macro-line/)

* SolidWorksMacro doesn't contain a definition for 'swApp' (VSTA)
  * [Solution 1](/solidworks-api/troubleshooting/macros/vsta-invalid-namespace/)

* Cannot open (for VSTA macros)
  * [Solution 1](/solidworks-api/troubleshooting/macros/run-vsta-macro-error/)

{% catalogue { type: sw-macro-fix } %}
