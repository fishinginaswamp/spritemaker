# Sprite maker spreadsheet

This was designed in google sheets, so I don't know how well this will actually work.

## Usage
Import the csv into google sheets

Open Apps Script and import the two ```.gs``` files

Certain sheets have different ways to generate code for sprites

Can be used for raw color format or with a color lookup table/palette

It will ask for permissions, grant them (you can always revoke them), run the function for the type of sprite you want. The code is pretty straightforward and you can always revoke access.

The output will be in the dev console, which you can copy/paste into a ```#include``` project file

This is "as-is" and it's up to the user to kinda figure out how it works and how to use it, sorry

Some of the sprites will automatically update via the onEdit() function. Others will have to be explicitly called from the Apps Script tab (the console is required)
