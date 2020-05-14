---
layout: sw-tool
title: Export part to Parasolid via Document Manager API (without SOLIDWORKS)
caption: Export Part File To Parasolid
description: Power Shell script to export part file to parasolid format (.xmp_bin) from command line via Document Manager API (without SOLIDWORKS)
image: /solidworks-document-manager-api/document/export-parasolid/export-parasolid.png
labels: [export,parasolid]
group: Import/Export
---
This PowerShell script allows exporting the SOLIDWORKS part file into the parasolid format (.xmp_bin) from the command line using SOLIDWORKS Document Manager API

This file can be opened in any compatible CAD application (SOLIDWORKS, Solid Edge, etc.)

This script doesn't require SOLIDWORKS to be installed and doesn't consume SOLIDWORKS license.

## Configuration and usage instructions

* Create two files and paste the code from the below snippets

### export-parasolid.ps1

{% code-snippet { file-name: export-parasolid.ps1 } %}

### export-parasolid.cmd

{% code-snippet { file-name: export-parasolid.cmd } %}

* Copy the *SolidWorks.Interop.swdocumentmgr.dll* into the folder where the above scripts are created. PowerShell script is based on .NET Framework 2.0 so the SOLIDWORKS Document Manager interop must target this framework. The dll can be found at: **SOLIDWORKS Installation Folder**\api\redist\CLR2\SolidWorks.Interop.swdocumentmgr.dll

Alternatively full path to interop can be specified as shown below. In this case it is not required to copy this dll into the folder with script files.

~~~ ps1
$Assem = ( 
   "Full path to SolidWorks.Interop.swdocumentmgr.dll"
    ) 
~~~

* Start the command line and execute the following command

~~~ bat
> [Full Path To export-parasolid.cmd] [Full Path To Input SOLIDWORKS file] [Full Path to output directory]
~~~

As the result all bodies from all configurations of the file are exported to the specified directory (directory is automatically created if not exist). Output files are named as following: *[original file name]_[configuration name].xmp_bin* The process log is displayed directly in the console:

![Parasolid export console output](export-parasolid-console-output.png)