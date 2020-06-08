---
title: Concatenate description variable on a data card button click using SOLIDWORKS PDM API
caption: Concatenate Description Variable
description: This example demonstrates how to handle the button click on data card and concatenate value of description variable based on values of revision and number variables in PDM add-in using SOLIDWORKS PDM Professional API
image: button-update-variable.gif
labels: [hook, button click, variables]
---
This example demonstrates how to handle the button click on data card using SOLIDWORKS PDM API and set the value of description variable based on concatenation of values of revision and number variables in PDM add-in.

* Add button to data card in PDM administration console
* Set the *Command type* option to *Run Add-in*
* Specify the name of the add-in as shown on the picture below. This option allows to specify the unique tag which can be retrieved by the add-in when button is clicked to properly identify the button.

![Data card setup options](data-card-button.png){ width=500 }

This option should be equal to the value specified in the *BUTTON_TAG* constant in the add-in

~~~ cs
private const string BUTTON_TAG = "_UpdateDesc_";
~~~

* Make sure that the *Number*, *Revision* and *Description* variables are present on the data card or modify the add-in code accordingly:

When button clicked the *Description* variable is updated by concatenation of values in *Number* and *Revision* variables.

![Description is updated when button is clicked](button-update-variable.gif){ width=450 }

{% code-snippet { file-name: add-in.cs } %}
