---
caption: Batch Export Models
title: Batch export SOLIDWORKS models via vbScript
description: Example of batch exporting SOLIDWORKS documents from the vbScript
---
This example of vbScript which demonstrates how to batch export SOLIDWORKS documents using vbScript

## Arguments

1. Path to folder with SOLIDWORKS models
1. Filter for the input files extension
1. Path to output folder
1. Extension of the output format

~~~
> "export-sw-models.vbs" "C:\Models" sldprt "C:\Output" step
~~~

{% code-snippet { file-name: export-sw-models.vbs } %}