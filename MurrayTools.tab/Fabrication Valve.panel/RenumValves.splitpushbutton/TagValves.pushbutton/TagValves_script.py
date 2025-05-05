import Autodesk
from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource, Transaction, FilteredElementCollector, Family, TransactionGroup,\
                                BuiltInCategory, FamilySymbol, BuiltInParameter, Reference, IndependentTag, TagMode, TagOrientation

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'Multi Category Tag - FP_Valve Number'
FamilyType = 'Multi Category Tag - FP_Valve Number'
# Check if the family is in the project
Fam_is_in_project = any(f.Name == FamilyName for f in families)

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

family_path = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES\Annotation\Multi Category Tag - FP_Valve Number.rfa'

tg = TransactionGroup(doc, "Add Valve Tags")
tg.Start()

t = Transaction(doc, 'Load Valve Tag Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_path, fload_handler)
t.Commit()

Pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework)\
                                                                 .WhereElementIsNotElementType()
familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation)\
                                                            .OfClass(FamilySymbol)\
                                                            .ToElements()
# Collect all existing tags in the current view
existing_tags = FilteredElementCollector(doc, curview.Id).OfClass(IndependentTag).ToElements()

ItmList1 = list()

t = Transaction(doc, 'Tag Valves')
t.Start()
for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        if not famtype.IsActive:
            famtype.Activate()
            doc.Regenerate()

for valve in Pipe_collector:
    ST = valve.ServiceType
    AL = valve.Alias
    if ST == 53 and AL not in ['STRAINER', 'CHECK', 'BALANCE']:
        is_tagged = False
        valve_id = valve.Id
        for tag in existing_tags:
            if tag.GetTaggedLocalElementIds == valve_id:
                is_tagged = True
                break
        
        # Only tag if not already tagged
        if not is_tagged:
            R = Reference(valve)
            ValveLocation = valve.Origin
            IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_MULTICATEGORY, TagOrientation.Horizontal, ValveLocation)

t.Commit()
tg.Assimilate()