
#Imports
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, Element, Transaction, TransactionGroup, FabricationPart
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsValueString

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Create a FilteredElementCollector to get all FabricationPart elements
AllElements = FilteredElementCollector(doc).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

t = Transaction(doc, "Update Line and Valve Numbers")
t.Start()

for Fpart in AllElements:
    #FP_Line Number
    ln = Fpart.GetPartCustomDataText(1)
    set_parameter_by_name(Fpart, 'FP_Line Number', ln)

    #FP_Valve Number
    vn = Fpart.GetPartCustomDataText(2)
    set_parameter_by_name(Fpart, 'FP_Valve Number', vn)
t.Commit()

