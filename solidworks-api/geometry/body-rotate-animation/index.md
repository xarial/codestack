---
layout: article
title: Create body rotation animation using SOLIDWORKS API
caption: Create Body Rotation Animation
description: VBA example to create a rotation animation of a selected body around Y axis using SOLIDWORKS API and temp bodies
image: body-rotate.gif
labels: [animation,rotate,temp body]
---
![Body rotation animation](body-rotate.gif)

This VBA example demonstrates how to create a rotation animation of a selected body in part document using SOLIDWORKS API.

There will be no additional features created in the Feature Manager tree. This macro **is not** using the SOLIDWORKS motion study. Body is rotated around Y axis at origin. Animation is created using the temp bodies and original body or feature manager tree is not affected.

Select body from the Feature Manager tree and run the macro.

![Body selected in the feature manager tree](feature-tree-body-selected.png){ width=250 }

Preview of the body is created and rotated until selection is cleared. When macro stops the original body is reverted to the original state.

{% code-snippet { file-name: Macro.vba } %}
