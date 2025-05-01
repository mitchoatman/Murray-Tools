#Imports
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, FabricationPart
from Parameters.Add_SharedParameters import Shared_Params

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
    if spool_param and spool_param.HasValue:
        custom_param = spool_param.AsString()
        if custom_param:
            Fpart.SpoolName = custom_param

t.Commit()
