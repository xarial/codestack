---
layout: article
title: Read and display body from the file using SOLIDWORKS API
caption: Read Body From File
description: VBA example to deserialize body geometry from external binary file into temp body and display using SOLIDWORKS API
image: /images/codestack-snippet.png
labels: [deserialize,com stream,temp body]
---
This VBA example demonstrates how to read the body geometry data from the external binary file. Load this data into the COM Stream and restore into the temp solid body using SOLIDWORKS API.

Body is displayed to the user and macro execution stops. Body is not present in the Feature Manager Tree and only visible in the graphics area.

Continue the macro execution to destroy the body.

{% code-snippet { file-name: Macro.vba } %}