---
layout: article
title: Sorting data in Visual Basic 6 (VBA)
caption: Sorting
description: Code example of various method of sorting data in Visual Basic 6 (VBA)
labels: [sort,bubble,logical]
---
## Logical Bubble Sorting

[Logical or natural sorting](https://en.wikipedia.org/wiki/Natural_sort_order) is an alphabetic sorting except of the multi-digit numeric values which are sorted as a single number.

For example the following logical order ab1, ab2, ab3, ab12 would be different in alphabetical order: ab1, ab12, ab2, ab3. Logical order is used when files are ordered in the Files Explorer.

The following example is sorting the specified string array in the logical order using [bubble sorting](https://en.wikipedia.org/wiki/Bubble_sort) technique.

*asc* parameter specifies if the values should be sorted in ascending or descending order:

~~~ vb
Dim vSortedArr As Variant
Dim vInputArr as Variant 'array of strings
vSortedArr = BubbleSort(vInputArr, False) 'sorting in descending order
~~~

{% code-snippet { file-name: logical-bubble-sort.vba } %}