---
layout: sw-tool
title: Export SOLIDWORKS files using SOLIDWORKS API in shell script
caption: Export Files
description: Script allows exporting of the SOLIDWORKS file into the foreign format using command line
image: export-file-result-console.png
labels: [export, script]
group: Import/Export
---
This PowerShell script allows exporting the SOLIDWORKS file into the specified foreign format from the command line using SOLIDWORKS API

## Configuration and usage instructions

* Create two files and paste the code from the below snippets

### export-file.ps1
{% code-snippet { file-name: export-file.ps1 } %}

### export-file.cmd
{% code-snippet { file-name: export-file.cmd } %}

* Copy the *SolidWorks.Interop.sldworks.dll* into the folder where the above scripts are created. PowerShell script is based on .NET Framework 2.0 so the SOLIDWORKS interop must target this framework. The dll can be found at: **SOLIDWORKS Installation Folder**\api\redist\CLR2\SolidWorks.Interop.sldworks.dll

![Script data files in the folder](script-folder.png){ width=450 }

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

![Messages in console reporting the progress and the result of exporting](export-file-result-console.png){ width=450 }