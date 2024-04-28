
import Autodesk
from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource, Transaction, FilteredElementCollector, Family, TransactionGroup,\
                                BuiltInCategory, FamilySymbol, BuiltInParameter, Reference, IndependentTag, TagMode, TagOrientation
import os

path, filename = os.path.split(__file__)
NewFilename = '\Pipe Accessory - Trimble Sleeve Size Tag.rfa'

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'Pipe Accessory - Trimble Sleeve Size Tag'
FamilyType = 'Pipe Accessory Tag - SLV Tag'
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

family_pathCC = path + NewFilename

tg = TransactionGroup(doc, "Add Sleeve Tags")
tg.Start()

t = Transaction(doc, 'Load Sleeve Size Family')
#Start Transaction
t.Start()
if Fam_is_in_project == False:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()

familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessoryTags)\
                                                             .OfClass(FamilySymbol)\
                                                             .ToElements()

ItmList1 = list()

t = Transaction(doc, 'Tag Sleeves')
t.Start()
for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        if famtype.IsActive == False:
            famtype.Activate()
            doc.Regenerate()

# Create a FilteredElementCollector to get Generic category elements
accessory_models_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PipeAccessory)
# Filter elements by name
accessory_elements = [element for element in accessory_models_collector if "Metal Sleeve" in element.Name or "Plastic Sleeve" in element.Name or "Cast Iron Sleeve" in element.Name]

for e in accessory_elements:
    R = Reference(e)
    loc = e.Location.Point
    IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, loc)

t.Commit()
#End Transaction Group
tg.Assimilate()









