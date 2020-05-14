---
layout: article
title: Get the sheet body geometry type using SOLIDWORKS API
caption: Get The Sheet Body Geometry Type
description: Example identifies the type of the selected sheet body (open shell, internal shell, external shell)
image: /solidworks-api/geometry/sheet-body-geometry-type/face-shell-types.png
labels: [example, face, geometry, open geometry, shell, solidworks api, topology]
redirect_from:
  - /2018/03/solidworks-api-geometry-get-body-geometry-type.html
---
There are 3 types of faces in SOLIDWORKS bodies:  

* **Open Shell**. Faces from the sheet bodies which are together with connected faces do not form the closed geometry (for example planar face, while face of the shell cube or sphere won't be considered as open)
* **Internal Shell**. Faces in solid bodies which belong to the cavities.
* **External Shell**. Any other faces which do not belong to previous groups

{% include img.html src="face-shell-types.png" width=400 height=243 alt="Shell types of face" align="center" %}

The example below identifies the type of the selected sheet body using SOLIDWORKS API. If the body is of open geometry (contains open shell faces) or closed geometry (no open shell faces). The closed geometry sheet body can be converted to a solid body.  

{% code-snippet { file-name: SolidWorksMacro.cs } %}
