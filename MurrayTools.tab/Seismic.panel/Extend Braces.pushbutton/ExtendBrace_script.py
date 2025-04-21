# Existing imports and setup remain unchanged
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import os

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return True

# Wrap the selection in a try-except block
try:
    bracesel = uidoc.Selection.PickObjects(ObjectType.Element,
        CustomISelectionFilter("Structural Stiffeners"), "Select seismic braces")
    Braces = [doc.GetElement(elId) for elId in bracesel]
except Autodesk.Revit.Exceptions.OperationCanceledException:
    # Handle the case where the user cancels the selection
    # from pyrevit import forms
    # forms.alert("Selection was canceled. Please select at least one seismic brace to proceed.")
    # Exit the script gracefully
    import sys
    sys.exit()

if len(Braces) > 0:
    # Second selection filter (for hanger)
    class CustomISelectionFilter(ISelectionFilter):
        def __init__(self, nom_categorie):
            self.nom_categorie = nom_categorie
        def AllowElement(self, e):
            if e.Category.Name == self.nom_categorie:
                return True
            else:
                return False
        def AllowReference(self, ref, point):
            return True

    # Prompt user for hanger selection, also wrapped in try-except
    try:
        pipesel = uidoc.Selection.PickObject(ObjectType.Element,
            CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hanger to set brace height")
        hanger = doc.GetElement(pipesel.ElementId)
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        # from pyrevit import forms
        # forms.alert("Hanger selection was canceled. Please select a fabrication hanger to proceed.")
        import sys
        sys.exit()

    # Get rod elevation
    STName = hanger.GetRodInfo().RodCount
    STName1 = hanger.GetRodInfo()
    for n in range(STName):
        rodloc = STName1.GetRodEndPosition(n)
        valuenum = rodloc.Z

    # Define functions
    def set_parameter_by_name(element, parameterName, value):
        element.LookupParameter(parameterName).Set(value)

    def get_parameter_value_by_name(element, parameterName):
        return element.LookupParameter(parameterName).AsDouble()

    # Start transaction
    t = Transaction(doc, 'Extend Seismic Braces')
    t.Start()

    for Brace in Braces:
        # Write data to TOS Parameter
        set_parameter_by_name(Brace, "Top of Steel", valuenum)
        # Read brace angle
        BraceAngle = get_parameter_value_by_name(Brace, "BraceMainAngle")
        sinofangle = math.sin(BraceAngle)
        # Read brace elevation
        BraceElevation = get_parameter_value_by_name(Brace, 'Offset from Host')
        # Equation to get the new hypotenuse
        Height = ((valuenum - BraceElevation) - 0.2330)
        newhypotenus = ((Height / sinofangle) - 0.2290)
        # Write new brace length to parameter
        set_parameter_by_name(Brace, "BraceLength", newhypotenus)

    # End transaction
    t.Commit()
else:
    from pyrevit import forms
    forms.alert('At least one element must be selected.')