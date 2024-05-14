import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import re
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString, set_parameter_by_name

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class MySelectionFilter(ISelectionFilter):
    def __init__(self):
        pass
    def AllowElement(self, element):
        if element.Category.Name == 'Generic Models':
            return True
        else:
            return False
    def AllowReference(self, element):
        return False
def fraction_to_float(fraction):
    parts = fraction.split(' ')
    whole = float(parts[0])
    if len(parts) > 1:
        numerator, denominator = map(float, parts[1].split('/'))
        fraction_value = numerator / denominator
    else:
        fraction_value = 0.0
    return whole + fraction_value

# selection
selected_element = uidoc.Selection.PickObject(ObjectType.Element, 'Select a Fabrication Part')
element = doc.GetElement(selected_element.ElementId)

size = element.LookupParameter('Overall Size').AsString()
numeric_part = size.rstrip('"')
try:
    OAsize = (float(numeric_part) / 12)
except:
    # Convert the fractional part to a float
    OAsize = (float(fraction_to_float(numeric_part) / 12))
pointdescription = str(round(((OAsize + 0.083) * 12) * 2) / 2) + ' Wall Sleeve'


selection_filter = MySelectionFilter()
wallsleeve = uidoc.Selection.PickElementsByRectangle(selection_filter)

t = Transaction(doc, 'Set Sleeve Size')
#Start Transaction
t.Start()

for sleeve in wallsleeve:
    parameterC = sleeve.LookupParameter('TS_Point_Description')
    if parameterC:
        parameterC.Set(pointdescription)
    parameter = sleeve.LookupParameter('Pipe OD')
    if parameter is not None:
        parameter.Set(OAsize)
    else:
        print("Parameter 'Pipe OD' not found on the sleeve:", sleeve)
t.Commit()
