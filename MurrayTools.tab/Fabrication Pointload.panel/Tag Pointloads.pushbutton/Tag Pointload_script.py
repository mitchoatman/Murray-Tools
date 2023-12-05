
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
FamilyName = 'Fabrication Hanger - Pointload Tag'
FamilyType = 'POINTLOAD'
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

family_pathCC = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES\Annotation\Fabrication Hanger - Pointload Tag.rfa'

tg = TransactionGroup(doc, "Add Pointload Tags")
tg.Start()

t = Transaction(doc, 'Load PointLoad Family')
#Start Transaction
t.Start()
if Fam_is_in_project == False:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()

Hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers)\
                                                                  .WhereElementIsNotElementType()
familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangerTags)\
                                                             .OfClass(FamilySymbol)\
                                                             .ToElements()

ItmList1 = list()

t = Transaction(doc, 'Tag Pointloads')
t.Start()
for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        if famtype.IsActive == False:
            famtype.Activate()
            doc.Regenerate()

for e in Hanger_collector:
        R = Reference(e)
        STName = e.GetRodInfo().RodCount
        ItmList1.append(STName)
        STName1 = e.GetRodInfo()
        for n in range(STName):
            rodloc = STName1.GetRodEndPosition(n)
            IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, rodloc)
t.Commit()
#End Transaction Group
tg.Assimilate()









