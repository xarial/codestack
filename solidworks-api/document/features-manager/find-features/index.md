---
layout: article
title: Find features in the tree by type and/or name pattern using SOLIDWORKS API
caption: Find Features
description: VBA macro to find all or first feature in the Feature Manager tree which match specific feature type name or name pattern using SOLIDWORKS API
image: /solidworks-api/document/features-manager/find-features/feature-manager-tree.png
labels: [traverse features,name pattern,type name]
---
{% include img.html src="feature-manager-tree.png" width=250 alt="Feature Manager Tree" align="center" %}

This VBA macro allows to find features in the Feature Manager tree using SOLIDWORKS API.

Features can be found by specifying the type name and/or name pattern (with support of wildcards). Specify empty string for name or type name to ignore this filter.

## Examples

~~~vb
Dim swFeat As SldWorks.Feature
Set swFeat = GetFirstFeature(swModel, "WeldMemberFeat") 'return first feature of WeldMemberFeat type (i.e. Structural Member)
~~~

~~~vb
Dim swFeat As SldWorks.Feature
Set swFeat = GetFirstFeature(swModel, "", "Sk*") 'return first feature which name starts with Sk
~~~

~~~vb
Dim vFeats As Variant
vFeats = GetAllFeatures(swModel, "WeldMemberFeat") 'return all features of WeldMemberFeat type (i.e. Structural Members)
~~~

~~~vb
Dim vFeats As Variant
vFeats = GetAllFeatures(swModel, "", "Sk*")'return all features whose names starts with Sk
~~~

{% code-snippet { file-name: Macro.vba } %}
