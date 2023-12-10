
import Autodesk
from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource, Transaction, FilteredElementCollector, Family, TransactionGroup,\
                                BuiltInCategory, FamilySymbol, BuiltInParameter, Reference, IndependentTag, TagMode, TagOrientation
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import os

path, filename = os.path.split(__file__)
NewFilename = '\Fabrication Hanger - Pointload Tag.rfa'

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True


    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


# Set desired family name and type name:
FamilyName = 'Fabrication Hanger - Pointload Tag'
FamilyType = 'POINTLOAD'
# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set a name to check for
FamilyName = "Fabrication Hanger - Pointload Tag"
# Check if the family is in the project
Fam_is_in_project = any(f.Name == FamilyName for f in families)
#print("Family '{}' is in project: {}".format(FamilyName, is_in_project))

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return true

fhangers = uidoc.Selection.PickObjects(ObjectType.Element,
CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabriction Hangers")            
Hanger_collector = [doc.GetElement( elId ) for elId in fhangers]

familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangerTags)\
                                                             .OfClass(FamilySymbol)\
                                                             .ToElements()

family_pathCC = path + NewFilename

tg = TransactionGroup(doc, "Selected Pointload Tags")
tg.Start()

t = Transaction(doc, 'Load PointLoad Family')
#Start Transaction
t.Start()
if Fam_is_in_project == False:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()

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









