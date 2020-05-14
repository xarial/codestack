---
layout: article
title: Usage of assembly component transformation in SOLIDWORKS API
caption: Component Transformation In The Assembly
description: Example explains transformation of rotation and translation for components in the assembly
image: /solidworks-api/geometry/transformation/assembly-component-transformation/comp-translation.png
labels: [acos, angle, component, example, orientation, point, position, rotation, solidworks api, transform, translation, vector]
redirect_from:
  - /2018/03/component-transformation-in-assembly.html
---
SOLIDWORKS components are instances of models (parts or assemblies) in the another parent assembly. Component's position in its space is driven by its transformation (regardless if the component is constrained by mates or moved in the space by free drag-n-drop operation). Transformation consists of 3 components: translation, rotation and scale.

To get the transformation of the component use the [IComponent2::Transform2](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.icomponent2~transform2.html) SOLIDWORKS API property. The transform in this case represents the relation of the component origin coordinate systems to the root assembly origin coordinate system. It is not required to multiple the transform of sub-assemblies for its children components to get the total transformation of these components relative to root assembly.

## Translation Transformation

In the example below the component is moved in space Along X, Y and Z coordinates. The following example will calculate the new positions of the component's origin:

{% include img.html src="comp-translation.png" width=640 alt="Component translational transform" align="center" %}

{% include_relative translation.vba.codesnippet %}

The following line will be output to the Watch window as the result of running the macro on [this sample model](transform-translation.SLDASM):

> Along X: 75mm; Along Y: -50mm; Along Z: -100mm

## Rotation Transformation

Now let's rotate the component and try to find the rotation angles. This component is rotated in all directions. **<span style="color: red;">Red line</span>** below - is the X axis of the assembly, **<span style="color: lime;">Green line</span>** - Y axis, **<span style="color: blue;">Blue line</span>** - Z axis. New X, New Y and New Z - are orientation of the corresponding axes in the component and dimensions indicate the angles between those axes.

{% include img.html src="comp-rotation.png" width=640 alt="Component rotational transform" align="center" %}

{% include_relative rotation.vba.codesnippet %}

Running the code above will output the following results for [this sample model](transform-rotation.SLDASM):

> Angle between X axes: 110 deg

> Angle between Y axes: 66.74 deg

> Angle between Z axes: 75 deg

## Preserving Transformation State In Configurations

By default transformation state of the floating component in the configuration will be overridden by another configuration state in case of assembly modifications, such as new component addition, mate changes etc. This is different from the manual behavior when floating component's position will not be changed if another configuration modified.

To demonstrate the issue consider the following scenario:

* Download the [sample assembly](preserve-transform.zip) which has a single component
* There are 2 configurations in the assembly
  * Configuration **A** has the component position fully defined by mates
  * Configuration **B** has a floating component without any mates in the random position
* Run the following macro. Macro will align the corner of the component with the origin of the assembly in the Configuration B

{% include img.html src="aligned-component.png" alt="Component's corner aligned with the origin of the assembly" align="center" %}

* Macro will stop at several points. Read the comment indicating the state
* On the last step the transformation assigned to the floating component was overridden by the transformation in the Configuration A driven by mates.

{% include_relative preserve.vba.codesnippet %}

In order to preserve the transformation it is required to [fix](/solidworks-api/document/assembly/components/fix-float/) the component in the Configuration B.

* Uncomment the following line
* Close the assembly without saving and reopen it again

~~~ vb
'FixComponentInThisConfiguration swComp
~~~

to

~~~ vb
FixComponentInThisConfiguration swComp
~~~

* Run macro again. Now the transformation is preserved