
#Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol
from rpw.ui.forms import TextInput
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import os
import re


DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

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


class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return true

bracesel = uidoc.Selection.PickObjects(ObjectType.Element,
CustomISelectionFilter("Structural Stiffeners"), "Select seismic braces")            
Braces = [doc.GetElement( elId ) for elId in bracesel]

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ExtendSeiBrace.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    f = open((filepath), 'w')
    f.write('3-7')
    f.close()

if len(Braces) > 0:

    f = open((filepath), 'r')
    PrevInput = f.read()
    f.close()

    #This displays dialog
    value = TextInput('TOS Elevation:', default = PrevInput)
    valuenum = parse_elevation(value)

    f = open((filepath), 'w')
    f.write(value)
    f.close()

    #start of defining functions to use
    def set_parameter_by_name(element, parameterName, value):
        element.LookupParameter(parameterName).Set(value)

    def get_parameter_value_by_name(element, parameterName):
        return element.LookupParameter(parameterName).AsDouble()

    t = Transaction(doc, 'Extend Seismic Braces')
    #Start Transaction
    t.Start()

    for Brace in Braces:
        #writes data to TOS Parameter
        set_parameter_by_name(Brace,"Top of Steel", valuenum)
        #reads brace angle
        BraceAngle = get_parameter_value_by_name(Brace, "BraceMainAngle")
        sinofangle = math.sin(BraceAngle)
        if RevitINT > 2019:
            #reads brace elevation
            BraceElevation = get_parameter_value_by_name(Brace, 'Offset from Host')
        else:
            BraceElevation = get_parameter_value_by_name(Brace, "Offset")
        #Equation to get the new hypotenus
        Height = ((valuenum - BraceElevation) - 0.2330)
        newhypotenus = ((Height / sinofangle) - 0.2290)
        #writes new Brace length to parameter
        set_parameter_by_name(Brace,"BraceLength", newhypotenus)
    #print("Brace Angle is: {}".format(DegBraceAngle))
    #print("Brace Elevation is: {}".format(BraceElevation))
    #print("Brace Height is: {}".format(Height))
    #print("Sin of Angle is: {}".format(sinofangle))

    #End Transaction
    t.Commit()
else:
    from pyrevit import forms
    forms.alert('At least one element must be selected.')
