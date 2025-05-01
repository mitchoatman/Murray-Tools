import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import re
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

#start of defining functions to use
def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_numvalue_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsString()

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

pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")            
Fhangers = [doc.GetElement( elId ) for elId in pipesel]

override_color = DB.Color(0, 255, 255)
override_settings = DB.OverrideGraphicSettings()
override_settings.SetProjectionLineColor(override_color)
view = doc.ActiveView


# start a transaction to modify model
t = Transaction(doc, 'Toggle Beam Hanger')
t.Start()

for hanger in Fhangers:
    BHangerStatus = get_parameter_value_by_name(hanger, 'FP_Beam Hanger')

    if BHangerStatus == None:
        set_parameter_by_name(hanger, 'FP_Beam Hanger', 'Yes')
        view.SetElementOverrides(hanger.Id, override_settings)
    elif BHangerStatus == 'No':
        set_parameter_by_name(hanger, 'FP_Beam Hanger', 'Yes')
        view.SetElementOverrides(hanger.Id, override_settings)
    else:
        set_parameter_by_name(hanger, 'FP_Beam Hanger', 'No')
        view.SetElementOverrides(hanger.Id, DB.OverrideGraphicSettings())

# end transaction
t.Commit()