
#Imports
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, Element, Transaction, TransactionGroup, FabricationPart

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Create a FilteredElementCollector to get all FabricationPart elements
AllElements = FilteredElementCollector(doc).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

#This writes to fab part custom data field
def set_customdata_by_custid(fabpart, custid, value):
	fabpart.SetPartCustomDataText(custid, value)

t = Transaction(doc, "Update Line and Valve Numbers")
t.Start()

for Fpart in AllElements:
    custom_param1 = Fpart.LookupParameter('FP_Line Number')
    if custom_param1.HasValue:
        custom_param = custom_param1.AsString()
        set_customdata_by_custid(Fpart, 1, custom_param)

    custom_param2 = Fpart.LookupParameter('FP_Valve Number')
    if custom_param2.HasValue:
        custom_param = custom_param2.AsString()
        set_customdata_by_custid(Fpart, 2, custom_param)

    custom_param6 = Fpart.LookupParameter('FP_Bundle')
    if custom_param6.HasValue:
        custom_param = custom_param6.AsString()
        set_customdata_by_custid(Fpart, 6, custom_param)

    spool_param = Fpart.LookupParameter('STRATUS Assembly')
    if spool_param.HasValue:
        custom_param = spool_param.AsString()
        Fpart.SpoolName = custom_param

t.Commit()

