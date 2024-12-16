__title__ = 'Pipe\nPointload'
__doc__ = """Calculates the combined weight of selected Fabrication Pipes
and divides that weight across the selected Fabrication Hangers.
1. Run Command.
2. Select Fabrication Pipes you wish to collect weight from.
3. Select Fabrication Hangers you wish to distribute the collected weight across.

Planned Improvements:
Add Functionality for Trapeze Hangers
Add Tagging functions
"""
__highlight__ = 'new'


import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import re
from SharedParam.Add_Parameters import Shared_Params

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

