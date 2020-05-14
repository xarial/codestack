---
layout: article
title: Get files paths from folder using Visual Basic 6 (VBA)
caption: Get Files From Folder
description: Function to get the list of all files in the folder with an option to traverse sub directories and specify the file extension using Visual Basic 6 (VBA)
image: /images/codestack-snippet.png
labels: [files,extension,traverse,recursive]
---
This function is Visual Basic 6 (VBA) allows to find the paths of files in the specified folder with an option to traverse sub directores and specifying the extension of files to return:

~~~ vb
vFiles = GetFiles("D:\MyFolder") 'get all files from the MyFolder directory in the D drive and all the sub folders
vFiles = GetFiles("D:\MyFolder", False) 'get only top level files from the MyFolder directory in the D drive
vFiles = GetFiles("D:\MyFolder", True, "txt") 'get all files in .txt format from the MyFolder directory in the D drive
~~~

{% code-snippet { file-name: GetFilesFromFolder.vba } %}