__title__ = 'Extend\nBraces'
__doc__ = """Extends Seismic braces to a user specified elevation.
1. Run Command.
2. Select braces you want to extend.
3. Enter elevation into dialog using format FT-IN (no spaces)
eg:  15-3 or 15-6.5
"""


#Imports
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

if len(Braces) > 0:
    
    #selection filter
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
    #prompt user for selection
    pipesel = uidoc.Selection.PickObject(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hanger to set brace height")            
    hanger = doc.GetElement(pipesel.ElementId)
    #look at selected hanger rod and get rod elevation Z
    STName = hanger.GetRodInfo().RodCount
    STName1 = hanger.GetRodInfo()
    for n in range(STName):
        rodloc = STName1.GetRodEndPosition(n)
        valuenum = rodloc.Z

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
        #reads brace elevation
        BraceElevation = get_parameter_value_by_name(Brace, 'Offset from Host')
        #Equation to get the new hypotenus
        Height = ((valuenum - BraceElevation) - 0.2330)
        newhypotenus = ((Height / sinofangle) - 0.2290)
        #writes new Brace length to parameter
        set_parameter_by_name(Brace,"BraceLength", newhypotenus)

    #End Transaction
    t.Commit()
else:
    from pyrevit import forms
    forms.alert('At least one element must be selected.')
