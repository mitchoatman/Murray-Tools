
#Imports
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, Element, Transaction, TransactionGroup, FabricationPart, Group
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

from Autodesk.Revit.DB import Group, Transaction

# Get group members and modify their parameters
group_elements = [Fpart.GetMemberIds() if isinstance(Fpart, Group) else [Fpart.Id] for Fpart in AllElements]

for group_member_ids in group_elements:
    for member_id in group_member_ids:
        member_element = doc.GetElement(member_id)  # Assuming you have the doc object
        ln = member_element.GetPartCustomDataText(1)
        set_parameter_by_name(member_element, 'FP_Line Number', ln)

        vn = member_element.GetPartCustomDataText(2)
        set_parameter_by_name(member_element, 'FP_Valve Number', vn)

        sn = member_element.SpoolName
        set_parameter_by_name(member_element, 'FP_Spool Number', sn)

  
t.Commit()

