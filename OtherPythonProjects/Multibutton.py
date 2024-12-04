from easygui import msgbox, multenterbox

fieldNames = ["Name", "Street Address", "City", "State", "ZipCode"]
fieldValues = list(multenterbox(msg='Fill in values for the fields.', title='Enter', fields=(fieldNames)))
#msgbox(msg=(fieldValues), title = "Results")
print(fieldValues[0])