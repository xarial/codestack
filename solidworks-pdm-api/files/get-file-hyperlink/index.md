---
layout: article
title: Get hyperlink to a file in SOLIDWORKS PDM vault (conisio url)
caption: Get File Hyperlink
description: PowerShell scripts which allows to get the consistent hyperlink (conisio url) to a specified file using PDM Professional API
image: hyperlink-email.png
labels: [conisio, hyperlink]
---
This PowerShell script allows extracting the conisio url to the specified file in the vault. This link can be used to get a persistent link to a file which can be used by any SOLIDWORKS PDM users.

SOLIDWORKS PDM API is used to extract the data required to form the conisio url: file id, folder id, etc.

Create 2 script files and paste the code below:

### get-url.ps1
{% code-snippet { file-name: get-url.ps1 } %}

### get-url.cmd
{% code-snippet { file-name: get-url.cmd } %}

Call the command line with the following parameters

* Vault Name
* Full path to a file
* Action for the hyperlink. Select one of the following: 
    * open
    * view
    * explore
    * get
    * lock
    * properties
    * history

For example:

~~~ cmd
get-url myvault "D:\myvault\part.sldprt" explore
~~~

The hyperlink is output to console:

![Conisio url is output to console window](conisio-url.png){ width=450 }

This hyperlink can be used now to access the file.

![Conisio url inserted to the link in e-mail](hyperlink-email.png){ width=450 }
