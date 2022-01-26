---
layout: sw-tool
title: Macro to force rebuild SOLIDWORKS document
caption: Rebuild Model
description: VBA macro to force rebuild, hide all types and show isometric view of SOLIDWORKS model
image: force-rebuild.svg
labels: [api, upgrade, performance, rebuild]
group: Performance
---
This VBA macro allows to perform operations usually required to upgrade the model to new version of SOLIDWORKS. It allows to:

* Force rebuild the model (ctrl+Q)

![Rebuild all configurations](rebuild-all-configurations.png)

* Set model to isometric orientation

![Zoom to fit](zoom-to-fit.png)

* Hide all view types

![Hide all types](view-hide-all-types.png)

Configure the macro actions by setting the values of corresponding constants

~~~ vb
Const DEFAULT_VIEWZOOMTOFIT As Boolean = True
Const DEFAULT_REBUILD As Boolean = True
Const DEFAULT_HIDE_ALL_TYPES As Boolean = True
~~~

This macro also supports [macro arguments](https://cadplus.xarial.com/macro-arguments/): **-zoomtofit**, **-rebuild**, **-hidealltypes**

![Macro arguments specified in Batch+](batch-plus-arguments.png)

{% code-snippet { file-name: Macro.vba } %}
