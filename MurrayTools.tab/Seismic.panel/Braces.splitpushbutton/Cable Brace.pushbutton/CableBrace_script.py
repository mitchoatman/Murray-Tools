
__title__ = 'Cable\nBrace'
__doc__ = """Inserts a Cable Seismic Brace Family.

1. Run Script.
2. Select where you would like to place family.
"""


import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'CABLE SEISMIC BRACE'
FamilyType = 'CABLE SEISMIC BRACE'
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

family_pathCC = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\Families\Structural Stiffeners (Seismic)\CABLE SEISMIC BRACE.rfa'

t = Transaction(doc, 'Load Cable Brace Family')
#Start Transaction
t.Start()
if Fam_is_in_project == False:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()


#Family symbol name to place.
symbName = 'CABLE SEISMIC BRACE'

#create a filtered element collector set to Category OST_Mass and Class FamilySymbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_StructuralStiffener)
collector.OfClass(FamilySymbol)

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
