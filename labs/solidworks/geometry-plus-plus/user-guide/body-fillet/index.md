---
layout: article
title: Bodies fillet feature in Geometry++
caption: Bodies Fillet
description: Feature allows adding the fillet to bodies, faces and edges and supports multi-bodies in SOLIDWORKS model
image: icon.png
toc-group-name: labs-solidworks-geometry-plus-plus
---
This commands allows adding the simple fillet to bodies, faces and edges. This command supports multi-bodies which means that single feature can be used to add fillets to different bodies.

![Bodies fillet property manager page](solid-bodies-fillet.png){ width=250 }

### Adding fillet to bodies

![Fillet added to a solid body](full-body-fillet.png){ width=250 }

Select the entire body from either from the feature tree or using the selection filter. Fillet will be added to each edge of the body.

### Adding fillet to faces

![Fillet added to face](face-fillet.png){ width=250 }

Select face or faces to add fillet to. Fillet is added to all edges of this face

### Adding fillet to edges

![Fillet added to edge](edge-fillet.png){ width=250 }

Select edge or edges to add fillet to.

### Adding fillet to vertices

![Fillet added to edges of vertex](vertex-fillet.png){ width=250 }

Select vertex or vertices to add fillet to. Fillet is added to all connected edges of this vertex.
