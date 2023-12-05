
#Imports
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, FabricationPart


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# Create a FilteredElementCollector to get all FabricationPart elements
AllElements = FilteredElementCollector(doc).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

t = Transaction(doc, 'Pin All Spools')
#Start Transaction
t.Start()

for i in AllElements:
    param_exist = i.LookupParameter("STRATUS Assembly")
    if param_exist.HasValue:
        i.Pinned = True

#End Transaction
t.Commit()
