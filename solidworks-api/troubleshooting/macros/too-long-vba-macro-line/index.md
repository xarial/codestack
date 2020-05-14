---
layout: issue-fix
title: Fix too long SOLIDWORKS VBA macro line error
caption: Too Long VBA Macro Line
description: Fixing the Compile error - Invalid outside procedure error when placing the long text into the VBA macro variable
image: /solidworks-api/troubleshooting/macros/too-long-vba-macro-line/doc-mgr-key-too-long.gif
labels: [macro, troubleshooting]
redirect-from:
  - /2018/04/macro-troubleshooting-too-long-vba-macro-line.html
---
## Symptoms

* SOLIDWORKS VBA macro is utilizing Document Manager APIs and new license was generated.
When generated license is placed into the macro some text highlighted red and *Compile error: Invalid outside procedure error* is displayed
* Macro is inserting static text into the note or custom properties. Text is replaced with new long text. Inserted string is highlighted and macro doesn't run

![Copy-pasting the Document Manager license key into the macro constant](doc-mgr-key-too-long.gif)

## Cause

Maximum number of symbols in a single line of VBA code is 1023.
It is not possible to insert more symbols without explicitly splitting the lines.
Pasting the line longer than the limit from the buffer will cause compilation errors.  

## Resolution

Split the line into multiple lines (no longer than 1023 symbols in single line) and use "string1" & _ "string2" to concatenate the lines.  

{% code-snippet { file-name: connect-to-dm-long-key.vba } %}
