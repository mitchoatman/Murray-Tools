import clr
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import re
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsValueString
import sys
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

fraction_pattern = re.compile(r"^(?P<num>[0-9]+)/(?P<den>[0-9]+)$")

# Selection filter class
class PickByCategorySelectionFilter(ISelectionFilter):
    """Selection filter for given categories"""
    def __init__(self, category_ids):
        self.category_ids = category_ids

    def AllowElement(self, element):
        if element.Category and element.Category.Id in self.category_ids:
            return True
        return False

    def AllowReference(self, reference, point):
        return False


def select_fabrication_elements():
    try:
        category_ids = [
            DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_FabricationDuctwork).Id,
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

# --- Parameter helpers ---
def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()

def round_up(n, decimals=0):
    multiplier = 10 ** decimals
    return math.ceil(n * multiplier) / multiplier
# --------------------------

# Collect user selections
Fpipework, Fhangers = select_fabrication_elements()

if len(Fpipework) > 0:
    # Iterate over fabrication ducts and collect weight data
    Total_Weight = 0.0
    for pipe in Fpipework:
        weight_str = get_parameter_value_by_name(pipe, 'Weight')
        if weight_str and isinstance(weight_str, str) and " lbm" in weight_str:
            lb_param = float(weight_str.replace(" lbm", ""))
            Total_Weight += lb_param

    if len(Fhangers) > 0:
        Hanger_Count = float(len(Fhangers))
        pointload = ((Total_Weight / Hanger_Count) / 10)

        t = Transaction(doc, 'Write Pointload Info')
        t.Start()

        for whanger in Fhangers:
            numofrods = whanger.GetRodInfo().RodCount
            if numofrods > 0:
                roundedpointload = pointload / numofrods
                set_parameter_by_name(whanger, "FP_Pointload", roundedpointload)

        t.Commit()
    else:
        TaskDialog.Show("Error", "At least one fabrication hanger must be selected.")
else:
    TaskDialog.Show("Error", "At least one fabrication duct must be selected.")