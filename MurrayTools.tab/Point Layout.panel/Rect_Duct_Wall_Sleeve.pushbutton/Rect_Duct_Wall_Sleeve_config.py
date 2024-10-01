import os
from rpw.ui.forms import FlexForm, TextBox, Button
from pyrevit import script

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Duct-Wall-Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

with open(filepath, 'r') as f:
    AnnularSpace = f.read()
    AnnularSpace = float(AnnularSpace) / 2

# Display dialog
components = [
    TextBox('space', str(AnnularSpace)),
    Button('Ok')
]
form = FlexForm('Sleeve Annular Space (Inches)', components)
form.show()

# Convert dialog input into variable
AnnularSpace = float(form.values['space'])
AnnularSpaceStored = str((AnnularSpace) * 2)

with open(filepath, 'w') as f:
    f.write (AnnularSpaceStored)