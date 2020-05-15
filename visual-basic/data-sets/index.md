---
layout: article
title: Data Sets (Arrays, Collections, Dictionaries) in Visual Basic
caption: Data Sets (Arrays, Collections, Dictionaries)
description: Article explaining the usage of arrays (one- and two-dimensional), collections (dynamic arrays), dictionaries in Visual Basic
order: 7
---
Data set is a list of elements stored within one variable. Data sets are used to hold family of values which are related to each other in certain way. For example data set may contain list of property names used in the document or a list of numbers of the average daily temperature within the year. Set may contain strongly typed data (i.e. set of String or set of Double) or it can contain set of variants.

Data sets of a fixed size are called arrays. The size is usually defined on the declaration, but arrays might also be resized. This is the most commonly used option to store the multiple data entries in Visual Basic.

Dynamic arrays with an ability to add new item on a fly without explicitly resizing it called Collections in Visual Basic or Lists (usually used in other programming languages). Collection usually exposes the property with a total number of elements and allows adding or removing elements from it as well as accessing specific item by index.

Array might contain the collection of Key-Value pairs and be indexed by Key. This enables several performance benefits when accessing element by key. These arrays are called Dictionary and usually only allow unique key in the collection.