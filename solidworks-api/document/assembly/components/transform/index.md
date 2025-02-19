---
caption: Transform Components
title: VBA macro to transform selected components in SOLIDWORKS assembly using the transformation
description: VBA macro to transform all selected components in the active SOLIDWORKS assembly using the predefined 4x4 transformation matrix
image: transform-component.svg
---

This macro can be useful for demo purposes where it is required to provide quick way of positioning the compnents in the assembly.

## Apply Transformation

This VBA macro applies the specified transformation to all selected components of the active SOLIDWORKS assembly.

Transformation fails if the component cannot be moved to the target position (e.g. has mates)

Transformation is 4x4 matrix specified in the **TRANSFORM** constant, represented as an array of 16 elements, separated by spaces.

Set the valu eof the **FIX_POSITION** to **True** to fix the component.

Use [Copy Transformation](#copy-transformation) macro to copy the transformation of the selected component into the clipboard.

{% code-snippet { file-name: Macro.vba } %}

## Copy Transformation

This macro copies the transformation of the selected component into the clipboard

{% code-snippet { file-name: CopyTransform.vba } %}

## Macro+

This macro supports [Macro+](https://cadplus.xarial.com/macro-plus/) arguments and transform can be passed to macro as first argument.

Use label icons below for the toolbar commands

[A](a.svg), [B](b.svg), [C](c.svg), [D](d.svg), [E](e.svg), [F](f.svg), [G](g.svg), [H](h.svg), [I](i.svg), [J](j.svg), [K](k.svg), [L](l.svg), [M](m.svg), [N](n.svg), [O](o.svg), [P](p.svg), [Q](q.svg), [R](r.svg), [S](s.svg), [T](t.svg), [Q](u.svg), [V](v.svg), [W](w.svg), [X](x.svg), [Y](y.svg), [Z](z.svg)