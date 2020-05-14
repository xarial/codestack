---
layout: article
title: Bodies fillet feature in Geometry++
caption: Bodies Fillet
description: Feature allows adding the fillet to bodies, faces and edges and supports multi-bodies in SOLIDWORKS model
image: /labs/solidworks/geometry-plus-plus/user-guide/body-fillet/icon.png
toc_group_name: labs-solidworks-geometry-plus-plus
---
This commands allows adding the simple fillet to bodies, faces and edges. This command supports multi-bodies which means that single feature can be used to add fillets to different bodies.

{% include img.html src="solid-bodies-fillet.png" width=250 alt="Bodies fillet property manager page" align="center" %}

### Adding fillet to bodies

{% include img.html src="full-body-fillet.png" width=250 alt="Fillet added to a solid body" align="center" %}

Select the entire body from either from the feature tree or using the selection filter. Fillet will be added to each edge of the body.

### Adding fillet to faces

{% include img.html src="face-fillet.png" width=250 alt="Fillet added to face" align="center" %}

Select face or faces to add fillet to. Fillet is added to all edges of this face

### Adding fillet to edges

{% include img.html src="edge-fillet.png" width=250 alt="Fillet added to edge" align="center" %}

Select edge or edges to add fillet to.

### Adding fillet to vertices

{% include img.html src="vertex-fillet.png" width=250 alt="Fillet added to edges of vertex" align="center" %}

Select vertex or vertices to add fillet to. Fillet is added to all connected edges of this vertex.
