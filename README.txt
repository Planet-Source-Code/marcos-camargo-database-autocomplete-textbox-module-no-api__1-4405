AutoComplete.bas
By MauTheMan - 1999

This module transforms any textbox to an AutoComplete Textbox.
You start typing and the module searches for a record that start
with the letters you typed, if it finds one, it completes the rest
of the textbox with the record.

HOW TO USE IT:

Simply add the module "AutoComplete.bas" to your project, call the
function CheckIsDelOrBack from the textbox keydown event (if you want
to have more than one AutoComplete textbox in the same form, set the
form's keypreview property to true and call the function CheckIsDelOrBack
from the form keydown event instead) - the parameter of the function
has to be "keycode" (like this : CheckIsDelOrBack(KeyCode))
then  call the function autocomplete from the textbox change event.

Use it, distribute it and so on (I would be happy to get an E-mail from
people who use this module - statistics purpose only)

Any problem, E-mail me:   mautheman@yahoo.com


