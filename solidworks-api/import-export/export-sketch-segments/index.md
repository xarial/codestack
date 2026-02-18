---
caption: Export Selected Sketch Segments
title: VBA macro to export selected sketch segments to IGES
description: VBA macro exports selected sketch segments and sketches to IGES format as wires 
---

This VBA macro exports only the selected sketch segments and sketches to IGES format as wireframe geometry. It is useful for generating input files for CAM software.

You can select individual sketch segments or entire sketches (in this case, all segments from the sketch will be processed). The macro supports selecting multiple segments and sketches at the same time.

Selected sketches (both 2D and 3D are supported) are exported to IGES format as wire entities.

The output file is saved in the same folder as the source file using the same base name. If a file with the same name already exists, an index is appended to the file name to prevent overwriting.

## Algorithm

* A new sketch is created and all selected sketch segments are converted into this sketch
* All sketch relations are removed
* A new part file is created and the sketch is copied into it
* IGES export settings are configured for wireframe output
* The original IGES settings are restored
* The temporary part file is closed
* The temporary sketch is deleted

{% code-snippet { file-name: Macro.vba } %}