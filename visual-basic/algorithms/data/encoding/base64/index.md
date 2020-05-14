---
layout: article
title: Encoding and decoding data in Base64 string format in Visual Basic 6 (VBA)
caption: Base64 String
description: Encoding and decoding byte array into Base64 string format in Visual Basic 6 (VBA)
image: /images/codestack-snippet.png
labels: [base64,encoding,decoding]
---
Base64 string allow to hold the byte array data in the string format

## Encode

~~~vb
Dim arr(5) As Byte
arr(0) = 1: arr(1) = 5: arr(2) = 2
arr(3) = 21: arr(4) = 101: arr(5) = 51

Dim base64Str As String
base64Str = ConvertToBase64String(arr) 'AQUCFWUz
~~~

{% code-snippet { file-name: encode-base64.vba } %}

## Decode

~~~vb
Dim base64Str As String
base64Str = "AQUCFWUz"

dim vArr As Variant
vArr = Base64ToArray(base64Str) 'Byte array: 1, 5, 2, 21, 101, 51
~~~

{% code-snippet { file-name: decode-base64.vba } %}