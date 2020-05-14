---
layout: sw-tool
title: Export SOLIDWORKS files using SOLIDWORKS API in shell script
caption: Export Files
description: Script allows exporting of the SOLIDWORKS file into the foreign format using command line
image: /solidworks-api/getting-started/scripts/power-shell/export-file/export-file-result-console.png
labels: [export, script]
categories: sw-tools
group: Import/Export
---
This PowerShell script allows exporting the SOLIDWORKS file into the specified foreign format from the command line using SOLIDWORKS API

## Configuration and usage instructions

* Create two files and paste the code from the below snippets

### export-file.ps1
{% include_relative export-file.ps1.codesnippet %}

### export-file.cmd
{% include_relative export-file.cmd.codesnippet %}

* Copy the *SolidWorks.Interop.sldworks.dll* into the folder where the above scripts are created. PowerShell script is based on .NET Framework 2.0 so the SOLIDWORKS interop must target this framework. The dll can be found at: **SOLIDWORKS Installation Folder**\api\redist\CLR2\SolidWorks.Interop.sldworks.dll

{% include img.html src="script-folder.png" width=450 alt="Script data files in the folder" align="center" %}

Alternatively full path to SOLIDWORKS interop can be specified as shown below. In this case it is not required to copy this dll into the folder with script files.

~~~ ps1
$Assem = ( 
   "Full path to SolidWorks.Interop.sldworks.dll"
    ) 
~~~

* Start the command line and execute the following command

~~~ bat
[Full Path To export-file.cmd] [Full Path To Input SOLIDWORKS file] [Full Path to output file and extension]
~~~

As the result the file is exported and the process log is displayed directly in the console:

{% include img.html src="export-file-result-console.png" width=450 alt="Messages in console reporting the progress and the result of exporting" align="center" %}