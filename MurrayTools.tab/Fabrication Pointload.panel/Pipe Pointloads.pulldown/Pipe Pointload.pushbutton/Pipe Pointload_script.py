import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.UI import TaskDialog
import math
import re
import sys
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

fraction_pattern = re.compile(r"^(?P<num>[0-9]+)/(?P<den>[0-9]+)$")

# Selection filter for MEP Fabrication Pipework and Hangers
class PickByCategorySelectionFilter(ISelectionFilter):
    """Selection filter implementation"""
    def __init__(self, category_ids):
        self.category_ids = category_ids

    def AllowElement(self, element):
        """Is element allowed to be selected?"""
        if element.Category and element.Category.Id in self.category_ids:
            return True
        return False

    def AllowReference(self, reference, point):
        """Not used for selection"""
        return False

def select_fabrication_elements():
    try:
        category_ids = [
            DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_FabricationPipework).Id,
            DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_FabricationHangers).Id
        ]
        msfilter = PickByCategorySelectionFilter(category_ids)
        selection = uidoc.Selection.PickObjects(ObjectType.Element, msfilter, "Select MEP Fabrication Pipes and Hangers")
        pipes = []
        hangers = []
        for ref in selection:
            element = doc.GetElement(ref.ElementId)
            if element.Category and element.Category.Id == category_ids[0]:  # OST_FabricationPipework
                pipes.append(element)
            elif element.Category and element.Category.Id == category_ids[1]:  # OST_FabricationHangers
                hangers.append(element)
        return pipes, hangers
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        TaskDialog.Show("Selection Cancelled", "Selection was cancelled by the user. Please select at least one MEP Fabrication Pipe or Hanger to continue.")
        sys.exit()  
    except Exception, e:
        TaskDialog.Show("Error", "An unexpected error occurred during selection: " + str(e))
        sys.exit()  

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

Fpipework, Fhangers = select_fabrication_elements()

# Iterate over fabrication pipes and collect length data
Total_Weight = 0.0

for pipe in Fpipework:
    if pipe.ItemCustomId == 2041:
        piperad = (get_parameter_numvalue_by_name(pipe, 'Main Primary Diameter') * 6)
        pipelength = get_parameter_numvalue_by_name(pipe, 'Length')
        B = (piperad * piperad * 3.14159)
        C = ((pipelength * 12) * B)
        D = (C / 231)
        Z = (D * 8.34)
        
        # Safely handle Weight parameter
        pweight_param = get_parameter_value_by_name(pipe, 'Weight')
        pipelb_param = 0.0
        
        if pweight_param and isinstance(pweight_param, str) and " lbm" in pweight_param:
            try:
                pipelb_param = float(pweight_param.replace(" lbm", "").strip())
            except ValueError:
                pipelb_param = 0.0  # Fallback
        
        F = pipelb_param + Z
        Total_Weight += F
    else:
        fweight_param = get_parameter_value_by_name(pipe, 'Weight')
        if fweight_param and isinstance(fweight_param, str) and " lbm" in fweight_param:
            fittinglb_param = float(fweight_param.replace(" lbm", ""))
            Total_Weight = Total_Weight + fittinglb_param

if len(Fhangers) > 0:
    Hanger_Count = 0.0
    
    for hanger in Fhangers:
        hangercount = 1
        Hanger_Count = Hanger_Count + hangercount
    
    pointload = ((Total_Weight / Hanger_Count) / 10)

    t = Transaction(doc, 'Write Pointload Info')
    t.Start()

    for whanger in Fhangers:
        numofrods = whanger.GetRodInfo().RodCount
        if numofrods > 0:
            roundedpointload = round_up(pointload) / numofrods
            set_parameter_by_name(whanger, "FP_Pointload", roundedpointload)
    
    t.Commit()
else:
    TaskDialog.Show("Error", "At least one fabrication hanger must be selected.")