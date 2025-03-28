
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family
import os

path, filename = os.path.split(__file__)
NewFilename1 = '\Fabrication Duct - Alias Tag.rfa'

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True


    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'Fabrication Duct - Alias Tag'
FamilyType = 'Fabrication Duct - Alias Tag'
# Check if the family is in the project
Fam_is_in_project = any(f.Name == FamilyName for f in families)
#print("Family '{}' is in project: {}".format(FamilyName, is_in_project))

family_pathCC1 = path + NewFilename1

t = Transaction(doc, 'Load Duct Alias Family')
#Start Transaction
t.Start()
if Fam_is_in_project == False:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC1, fload_handler)
    family = doc.LoadFamily(family_pathCC2, fload_handler)
t.Commit()

#Family symbol name to place.
symbName = 'Fabrication Duct - Alias Tag'

#create a filtered element collector set to Category OST_Mass and Class FamilySymbol
collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationDuctworkTags).OfClass(FamilySymbol)

famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()

#Search Family Symbols in document.
for item in famtypeitr:
    famtypeID = item
    famsymb = doc.GetElement(famtypeID)

    #If the FamilySymbol is the name we are looking for, create a new instance.
    if famsymb.Family.Name == symbName:
        uidoc.PostRequestForElementTypePlacement(famsymb)

#t.Commit()
