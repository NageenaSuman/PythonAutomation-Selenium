import easygui
from easygui import *

# window title
title = "PDF Resolution I/P in DPI"

# Get user input
myvar = easygui.enterbox("Enter the value for DPI:", title)
myvar1 = easygui.enterbox("Enter the PDF File Path:")
myvar2 = easygui.enterbox("Enter the Output folder:")

# Now 'myvar' contains the user's input
print(f"Your favorite color is: {myvar}")
