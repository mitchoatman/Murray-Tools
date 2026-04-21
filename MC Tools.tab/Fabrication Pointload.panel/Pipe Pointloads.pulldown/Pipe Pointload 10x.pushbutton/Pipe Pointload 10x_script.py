import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import re
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView


#start of defining functions to use
def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_numvalue_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()

def round_up(n, decimals=0):
    multiplier = 10 ** decimals
    return math.ceil(n * multiplier) / multiplier    
#end of defining functions to use
hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

t = Transaction(doc, 'Divide Pointload 10x')
#Start Transaction
t.Start()
for hanger in hanger_collector:
    ploadvalue = get_parameter_numvalue_by_name(hanger, 'FP_Pointload')
    try:
        newploadvalue = round_up(ploadvalue / 10)
        if newploadvalue:
            set_parameter_by_name(hanger, 'FP_Pointload', newploadvalue)
    except:
        sys.exit() 
    
#End Transaction
t.Commit()

