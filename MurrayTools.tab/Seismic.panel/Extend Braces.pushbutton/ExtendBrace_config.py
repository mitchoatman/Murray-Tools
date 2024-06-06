__title__ = 'Extend\nBraces'
__doc__ = """Extends Seismic braces to a user specified elevation.
1. Run Script.
2. Select braces you want to extend.
3. Enter elevation into dialog using format FT-IN (no spaces)
eg:  15-3 or 15-6.5
"""


#Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol
from rpw.ui.forms import TextInput
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import os


DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)




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
    value = TextInput('TOS Elevation *Input in this format FT-IN*', default = PrevInput)
    InputFT = float(value.split("-", 1)[0])
    InputIN = (float(value.split("-", 1)[1]) / 12)
    #print("Sin of Angle is: {}".format(InputFT))
    #print("Sin of Angle is: {}".format(InputIN))
    valuenum = InputFT + InputIN

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
