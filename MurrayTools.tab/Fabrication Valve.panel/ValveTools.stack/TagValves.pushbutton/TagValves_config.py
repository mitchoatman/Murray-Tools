
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
FamilyName = 'Fabrication Pipe - Valve Number'
FamilyType = 'Fabrication Pipe - Valve Number'
# Check if the family is in the project
Fam_is_in_project = any(f.Name == FamilyName for f in families)
#print("Family '{}' is in project: {}".format(FamilyName, is_in_project))

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True


    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

family_pathCC1 = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES\Annotation\Fabrication Pipe - Valve Number.rfa'
family_pathCC2 = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES\Annotation\Fabrication Pipe - Valve Tag.rfa'

tg = TransactionGroup(doc, "Add Valve Tags")
tg.Start()

t = Transaction(doc, 'Load Valve Tag Family')
#Start Transaction
t.Start()
if Fam_is_in_project == False:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC1, fload_handler)
    family = doc.LoadFamily(family_pathCC2, fload_handler)
t.Commit()

Pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework)\
                                                                  .WhereElementIsNotElementType()
familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangerTags)\
                                                             .OfClass(FamilySymbol)\
                                                             .ToElements()

ItmList1 = list()

t = Transaction(doc, 'Tag Valves')
t.Start()
for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        if famtype.IsActive == False:
            famtype.Activate()
            doc.Regenerate()

for valve in Pipe_collector:
    ST = valve.ServiceType
    if ST == 53:
        R = Reference(valve)
        ValveLocation = valve.Origin
        IndependentTag.Create(doc, curview.Id, R, True, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, ValveLocation)
t.Commit()
#End Transaction Group
tg.Assimilate()









