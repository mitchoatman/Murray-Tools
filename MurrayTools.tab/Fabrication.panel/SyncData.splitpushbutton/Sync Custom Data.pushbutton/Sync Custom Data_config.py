#Imports
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, FabricationPart, Group
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsValueString
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Create a FilteredElementCollector to get all FabricationPart elements
AllElements = FilteredElementCollector(doc).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

#This writes to fab part custom data field
def set_customdata_by_custid(fabpart, custid, value):
	fabpart.SetPartCustomDataText(custid, value)

t = Transaction(doc, "Update Fabrication Custom Data")
t.Start()

# Dictionary of parameter names and corresponding custom data IDs
param_map = {
    'FP_Line Number': 1,
    'FP_Valve Number': 2,
    'FP_Bundle': 6,
    'FP_Location': 13
}
# Gets revit data and pushes into fabricaton part custom data fields
for Fpart in AllElements:
    # Retrieve all relevant parameters once
    param_values = {p: Fpart.LookupParameter(p) for p in param_map}
    
    # Iterate through only valid parameters
    for param_name, custom_id in param_map.items():
        param = param_values[param_name]
        if param and param.HasValue:
            custom_param = param.AsString()
            if custom_param:
                set_customdata_by_custid(Fpart, custom_id, custom_param)
    
    # Handle SpoolName separately
    spool_param = Fpart.LookupParameter('STRATUS Assembly')
    spool_param = Fpart.LookupParameter('FP_Spool Number')
    if spool_param:
        custom_param = spool_param.AsString()
        if custom_param:
            Fpart.SpoolName = custom_param

t.Commit()

# Pulls data from imported maj fab parts into revit parameters
t = Transaction(doc, "Update Line and Valve Numbers")
t.Start()
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