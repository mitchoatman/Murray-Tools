
#Imports
from Autodesk.Revit import DB
from Autodesk.Revit.DB import *
from rpw.ui.forms import TextInput
from Autodesk.Revit.UI.Selection import *
import math

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

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
Hanger = [doc.GetElement( elId ) for elId in pipesel]

Dimensions = []

def myround(x, multiple):
    return multiple * math.ceil(x/multiple)

bvalue_abvstd = 0.0

t = Transaction(doc, "Round Trapeze Width")
t.Start()

if Hanger:
    for e in Hanger:
        e.GetHostedInfo().DisconnectFromHost()
        for dim in e.GetDimensions():
            Dimensions.append(dim.Name)
            if dim.Name == "Width":
                width_value = e.GetDimensionValue(dim)
                # e.SetDimensionValue(dim, width_value)
            if dim.Name == "Bearer Extn":
                bearer_value = e.GetDimensionValue(dim)
                in_bvalue = (bearer_value * 12)
                if in_bvalue > 4.0:
                    bvalue_abvstd = in_bvalue - 4.0
                in_wvalue = (width_value * 12)
                rnd_value = myround((in_bvalue + in_wvalue + bvalue_abvstd), 2)
                abv_value = rnd_value - in_wvalue
                hlf_diff = (abv_value - 4.0) / 2
                new_value = (abv_value - hlf_diff) / 12
                e.SetDimensionValue(dim, new_value)
t.Commit()
