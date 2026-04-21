from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family
import os

path, filename = os.path.split(__file__)
NewFilename = '\Fabrication Pipe - Size Tag - Aligned.rfa'

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

# Set desired family name and type name
FamilyName = 'Fabrication Pipe - Size Tag - Aligned'
FamilyType = 'Size Tag'

family_pathCC = path + NewFilename

# Check if we need to load the family first
families = FilteredElementCollector(doc).OfClass(Family)
needs_loading = not any(f.Name == FamilyName for f in families)

if needs_loading:
    t = Transaction(doc, 'Load Pipe Size Family')
    t.Start()
    try:
        fload_handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(family_pathCC, fload_handler)
        t.Commit()
    except Exception, e:
        print "Error loading families: %s" % str(e)
        t.RollBack()
        raise

# Now handle the symbol placement - no transaction needed for PostRequest
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_FabricationPipeworkTags)
collector.OfClass(FamilySymbol)

target_symbol = None
for symbol in collector:
    if symbol.Family.Name == FamilyName and symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType:
        target_symbol = symbol
        break

if target_symbol:
    if not target_symbol.IsActive:
        t = Transaction(doc, 'Activate Family Symbol')
        t.Start()
        try:
            target_symbol.Activate()
            t.Commit()
        except Exception, e:
            print "Error activating symbol: %s" % str(e)
            t.RollBack()
            raise
    
    uidoc.PostRequestForElementTypePlacement(target_symbol)
else:
    print "Could not find family type 'Size Tag' in family 'Fabrication Pipe - Size Tag'"