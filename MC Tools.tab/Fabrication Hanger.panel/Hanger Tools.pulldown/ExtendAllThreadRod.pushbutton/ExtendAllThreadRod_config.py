# Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from rpw.ui.forms import FlexForm, Label, TextBox, Button, CheckBox
from pyrevit import script
import os

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, FamilySymbol, Structure, Transaction,
    BuiltInParameter, Family, TransactionGroup, FamilyInstance
)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, category_name, target_name):
        self.category_name = category_name
        self.target_name = target_name

    def AllowElement(self, e):
        if e.Category.Name != self.category_name:
            return False
        name_param = e.LookupParameter("Family")
        if name_param and name_param.AsValueString() == self.target_name:
            return True
        return False

    def AllowReference(self, ref, point):
        return True

try:
    pipesel = uidoc.Selection.PickObjects(
        ObjectType.Element,
        CustomISelectionFilter("MEP Fabrication Pipework", "All Thread Rod"),
        "Select 'All Thread Rod' Elements Only"
    )
    Hanger = [doc.GetElement(elId) for elId in pipesel]
except Autodesk.Revit.Exceptions.OperationCanceledException:
    TaskDialog.Show("Selection Cancelled", "Selection Cancelled by User.")
    import sys
    sys.exit()

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ExtendHangerRod.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('5-6')

with open(filepath, 'r') as f:
    PrevInput = f.read()

if len(Hanger) > 0:
    # Display dialog
    components = [
        Label('TOS Elevation from 0 *Input in this format FT-IN*'),
        TextBox('Elevation', PrevInput),
        Button('Ok')
    ]
    form = FlexForm('Modify All Thread Rod', components)
    form.show()

    # Convert dialog input into variable
    try:
        value = form.values['Elevation']
        InputFT = float(value.split("-", 1)[0])
        InputIN = float(value.split("-", 1)[1]) / 12
        valuenum = InputFT + InputIN
    except (KeyError, IndexError, ValueError):
        TaskDialog.Show("Input Error", "Invalid elevation input. Please use FT-IN format (e.g., 5-6).")
        import sys
        sys.exit()

    with open(filepath, 'w') as f:
        f.write(value)

    t = Transaction(doc, 'Extend All Thread Rods')
    t.Start()

    for e in Hanger:
        bbox = e.get_BoundingBox(None)
        if bbox:
            top_z = bbox.Max.Z
            delta = valuenum - top_z
            length_param = e.LookupParameter("Length")
            if length_param and not length_param.IsReadOnly:
                current_length = length_param.AsDouble()
                length_param.Set(current_length + delta)
            else:
                output = script.get_output()
                print("Length parameter not found or read-only on: {}".format(output.linkify(e.Id)))
        else:
            output = script.get_output()
            print("Bounding box not found for element: {}".format(output.linkify(e.Id)))

    t.Commit()
else:
    TaskDialog.Show("Selection Error", "At least one 'All Thread Rod' element must be selected.")