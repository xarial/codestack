---
layout: article
title: Power Shell script to set folder permissions using SOLIDWORKS PDM API
caption: Set Folder Permissions
description: Vb.NET Power Shell script to set permissions to specified folder for specified user using SOLIDWORKS PDM API
lang: en
image: /solidworks-pdm-api/permissions/set-folder-permissions/folder-permissions.png
labels: [permissions,folder]
---
{% include img.html src="folder-permissions.png" width=450 alt="Folder permissions in SOLIDWORKS PDM Administration panel" align="center" %}

This Power shell script allows to set the specified folder permissions for the specified user using SOLIDWORKS PDM API.

To use script create PowerShell file and command line file as shown below.

It is required to place the SOLIDWORKS PDM interop into the same folder as script files. Refer [Interops in .NET for Framework 2.0](/solidworks-pdm-api/getting-started#framework-20-or-older) article for more information about generating the interop.

Scrip arguments:

1. *vaultName* - name of the vault to perform the operation
1. *userName* - user name who performs the operation (should have a permission to assign permissions)
1. *password* - password for the user name above
1. *folderName* - full path to folder to change permission for
1. *targetUserName* - user name to change permissions for
1. *permissions* - permissions to assign. Integer number which represents a single permissions or a group of permissions. Permissions numbers defined in [EdmRightFlags](http://help.solidworks.com/2018/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.EdmRightFlags.html). Sum the values of required permissions to assign multiple values (e.g. set 1 for read files permission and 15 for read, check out, delete and add files [1 + 2 + 4 + 8])

~~~
> set-permissions.cmd MyVault admin pwd "D:\My Vaults\Vault1\Folder1" user1 15
~~~

## set-permissions.ps1

{% include_relative set-permissions.ps1.codesnippet %}

## set-permissions.cmd

{% include_relative set-permissions.cmd.codesnippet %}
