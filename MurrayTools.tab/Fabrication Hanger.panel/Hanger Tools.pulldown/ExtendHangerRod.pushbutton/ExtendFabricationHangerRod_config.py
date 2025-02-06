# Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from rpw.ui.forms import FlexForm, Label, TextBox, Separator, Button, CheckBox
import re
import os

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, category_name):
        self.category_name = category_name

    def AllowElement(self, e):
        return e.Category.Name == self.category_name

    def AllowReference(self, ref, point):
        return True

def parse_elevation(input_str):
    """
    Converts various user-input formats into feet as a float.
    Supported formats:
    - 5-6         -> 5 feet 6 inches
    - 5 6         -> 5 feet 6 inches
    - 5.5         -> 5.5 feet
    - 5'-6"       -> 5 feet 6 inches
    - 5' 6"       -> 5 feet 6 inches
    - 5'-6 3/8"   -> 5 feet 6.375 inches
    - 5'-6.125"   -> 5 feet 6.125 inches
    """
    input_str = input_str.strip().replace('"', '')  # Remove quotes if present

    # Case 1: Decimal feet (e.g., "5.5")
    if re.match(r"^\d+(\.\d+)?$", input_str):
        return float(input_str)

    # Case 2: Feet-inches format (handles "5-6", "5 6", "5'-6 3/8", "5'-6.125")
    match = re.match(r"(\d+)[\s'\-]*(\d*(?:\s*\d+/\d+|\.\d+)*)?", input_str)
    if match:
        feet = float(match.group(1))
        inches = 0

        if match.group(2):
            inch_part = match.group(2).strip()
            if " " in inch_part:  # Handles mixed whole + fraction ("6 3/8")
                whole_inches, fraction = inch_part.split(" ", 1)
                inches = float(whole_inches) + eval(fraction)
            elif "/" in inch_part:  # Handles fraction-only inches ("3/8")
                inches = eval(inch_part)
            elif "." in inch_part:  # Handles decimal inches ("6.125")
                inches = float(inch_part)
            elif inch_part.isdigit():  # Handles whole inches ("6")
                inches = float(inch_part)

        return feet + (inches / 12)

    raise ValueError("Invalid elevation format. Use 5-6, 5.5, 5'-6\", 5'-6.125\", etc.")

# Select fabrication hangers
pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers to Extend")
Hangers = [doc.GetElement(elId) for elId in pipesel]

# Ensure folder for settings
folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ExtendHangerRod.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('5-6')

with open(filepath, 'r') as f:
    PrevInput = f.read()

if len(Hangers) > 0:
    # Display user input dialog
    components = [
        Label('Enter Target Elevation:'),
        TextBox('Elevation', PrevInput),
        CheckBox('RodControl', 'Add Rod Control', default=True),
        Button('Ok')
    ]
    form = FlexForm('Modify Hanger Rod', components)
    form.show()

    # Parse user input
    value = form.values['Elevation']
    try:
        target_elevation = parse_elevation(value)
    except ValueError as e:
        print(e)
        exit()
    RodControl = form.values['RodControl']

    with open(filepath, 'w') as f:
        f.write(value)

    t = Transaction(doc, 'Extend Hanger Rods')
    t.Start()

    for hanger in Hangers:
        rod_info = hanger.GetRodInfo()
        rod_count = rod_info.RodCount

        # Get the hanger's associated level
        level_id = hanger.get_Parameter(DB.BuiltInParameter.FABRICATION_LEVEL_PARAM).AsElementId()
        level = doc.GetElement(level_id)
        level_elevation = level.Elevation if level else 0  # Get level elevation in feet

        # Normalize target elevation relative to level
        adjusted_elevation = target_elevation - level_elevation

        for n in range(rod_count):
            rod_length = rod_info.GetRodLength(n)
            rod_end_pos = rod_info.GetRodEndPosition(n)
            rod_z = rod_end_pos.Z  # Extract rod Z-coordinate

            # Adjust rod length to match target elevation
            new_length = rod_length + (adjusted_elevation - rod_z)
            rod_info.SetRodLength(n, new_length)

    t.Commit()
else:
    print('At least one fabrication hanger must be selected.')
