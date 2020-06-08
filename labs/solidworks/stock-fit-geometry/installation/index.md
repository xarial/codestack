---
title: Installation Guide for Stock Master add-ins for SOLIDWORKS
caption: Installation Guide
description: Instructions of installing of Stock Master add-in for SOLIDWORKS which provides additional features for packaging and stocking
toc-group-name: labs-solidworks-stock-master
order: 2
---
To install add-in download the latest msi-installer (*StockMaster.msi*) at [this link v. 0.5.0 (beta 1)](https://github.com/codestackdev/stock-fit-geometry/releases/tag/beta1).

## Note
If older version of the add-in (v. 0.0.3 or older) is installed it is required to manually uninstall the previous version by following steps below:

* Navigate to the installation folder of previous version (The path can be found by hovering the mouse over the *Stock Fit Geometry* addin in the Tools->Add-ins... menu in SOLIDWORKS)
* Run the following command from the command line (it might be required to run command line as administrator). Replace *FULL PATH TO CodeStack.StockFit.Sw.dll* with the corresponding full path to the dll.

~~~ bat
"%Windir%\Microsoft.NET\Framework64\v4.0.30319\regasm" /codebase /u "FULL PATH TO CodeStack.StockFit.Sw.dll"
~~~
* Delete the folder

This procedure is not required for any future versions of the add-in.