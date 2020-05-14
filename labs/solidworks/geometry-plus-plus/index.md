---
layout: article
title: Geometry++ - add-in which complements SOLIDWORKS geometry commands
caption: Geometry++
description: SOLIDWORKS add-in providing additional set of commands related to geometry modifications and creation
image: /labs/solidworks/geometry-plus-plus/logo.png
categories: sw-labs
group: Geometry
toc_group_name: labs-solidworks-geometry-plus-plus
---
{% include img.html src="logo.png" alt="Geometry++" align="center" %}

Geometry++ is a SOLIDWORKS add-in extending the functionality related to geometry creation and manipulation. Add-in is fully integrated into SOLIDWORKS and provides same look and feel of additional features as any other built-in features. All feature preserve parametric behavior and automatically regenerated when required.

Refer [Installation](installation) page for more information and download link.

Source code is available at [GitHub](https://github.com/codestackdev/geometry-plus-plus)

## Features

### Convert Solid To Surface

{% include img.html src="/labs/solidworks/geometry-plus-plus/user-guide/convert-solid-to-surface/icon.png" alt="Convert Solid To Surface" align="center" %}

This command converts solid bodies to surface bodies. Multiple input bodies can be converted within one feature.

### Crop Bodies

{% include img.html src="/labs/solidworks/geometry-plus-plus/user-guide/crop-bodies/icon.png" alt="Crop Bodies" align="center" %}

This command allows cropping surface and solid (target bodies) using sketches or sketch regions (trimming tools).

### Extrude Surface With Caps

{% include img.html src="/labs/solidworks/geometry-plus-plus/user-guide/extrude-surface-cap/icon.png" alt="Extrude Surface With Caps" align="center" %}

This command allows extruding the surface and adding the caps at the ends of extrusion.

### Bodies Fillet

{% include img.html src="/labs/solidworks/geometry-plus-plus/user-guide/body-fillet/icon.png" alt="Bodies Fillet" align="center" %}

This command allows adding the fillet to entire bodies, faces, edges and vertices supporting multiple bodies within a single feature

### Split Body By Faces

{% include img.html src="/labs/solidworks/geometry-plus-plus/user-guide/split-body-by-faces/icon.png" alt="Split Body By Faces" align="center" %}

This command allows creating of surface (sheet) bodies from all the faces of input solid or surface bodies.

## Performance

{% include img.html src="/labs/solidworks/geometry-plus-plus/user-guide/suspend-rebuild/icon.png" alt="Suspend Rebuild" align="center" %}

This command allows to temporary suspend the rebuild operation in SOLIDWORKS parts, assemblies and drawings allowing to combine multiple rebuild operations into one operation to reduce the rebuild time.