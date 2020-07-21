# Contributing Guidelines

[CodeStack](https://www.codestack.net) is built using the [Docify](https://docify.net) engine.

## Structure

*index.md* file represents a new article. This file should be placed into the folder which will represent the part of the url. This should be a short name.

Place all of the assets related to this article (e.g. images, code snippets, etc.) at the same folder next to the *index.md*

## Writing Content

Follow the [Creating static markdown content in Docify](https://docify.xarial.com/content/static/) article for the syntax of markdown used in CodeStack articles.

Currently, the only supported language is English, however, CodeStack is planning to extend its articles to other languages in the future.

Each page should contain the attribution defined in [page metadata](https://docify.xarial.com/metadata/)

~~~ yaml
---
caption: Short title of the page. This title is displayed in the table of content
title: Long title of the page. Contains the full information about the article
description: Detailed description of the article
image: Optional image associated with the article in .png or .svg format
---
~~~

## Referencing Code Snippets

It is recommended not to embed the code snippets directly into the article as it will make maintenance more complicated. Instead, use the syntax of [Code Snippet Plugin](https://docify.xarial.com/standard-library/plugins/code-snippet/).

Place the snippet file next to the article. Use the file extensions of the corresponding language. For VBA macros use .vba extension.

Refer the snippet in the article as follows:

~~~
{% code-snippet { file-name: Macro.vba } %}
~~~