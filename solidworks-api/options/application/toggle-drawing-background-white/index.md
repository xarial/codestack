---
layout: sw-tool
title:  VBA macro to toggle white background in drawings
caption: Toggle white drawing background
description: VBA macro to toggle a white background with another color of your choice in drawings using system settings
image: ToggleWhiteBackground-icon.svg
labels: [Drawings, options, background, capture]
group: Options 
---
Author: [Eddy Alleman](https://www.linkedin.com/in/eddyalleman/) ([EDAL Solutions](https://www.edalsolutions.be/index.php/en/))

![SolidWorks system options to set Drawing Background manually](solidworks-option-background.png){ width=450 }

Introduction
On the SolidWorks forum someone asked how to make a macro that toggles between the default drawing background color and a white color.
The goal was to make it easier to capture images where a white background was required.

Here is a simple macro that does exactly that. I will also explain the basic buttons/shortcuts/menus you need.

If you want to toggle between other colors, you can change that in the Color1 and Color2 constants below. 

## But how do we get that number that corresponds to the color we want?
Just change it to your favorite color manually in SolidWorks options (in the image above I choose for a more distinct yellow color)
Then open the macro with the macro editor (Menu Tools > Macro > Edit or use the Macros toolbar). 
Open the immediate window if it is not already visible (CTRL + G)
Run the macro (F5 or green arrow button) and in the immediate window you should see the color you choose represented by a number:

![Immediate Window showing the chosen color after running the macro](vba-immediate-window-chosen-color.png)

Adapt the number in the code (Color2) and when you run the macro the background color will switch between white and your favorite color.

{% code-snippet { file-name: Macro.vba } %}

