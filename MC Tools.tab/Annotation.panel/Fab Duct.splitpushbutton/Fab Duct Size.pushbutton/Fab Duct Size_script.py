import os
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (FilteredElementCollector, Transaction,
                               BuiltInCategory, FamilySymbol, Family)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

folder = os.path.dirname(__file__)
symbName = 'Fabrication Duct - Size Tag - Aligned'
fam_files = {
    'Fabrication Duct - Size Tag - Aligned':
        os.path.join(folder, 'Fabrication Duct - Size Tag - Aligned.rfa'),
    'Fabrication Duct - Size Tag':
        os.path.join(folder, 'Fabrication Duct - Size Tag.rfa'),
}

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

# Load only what is missing
loaded = set(f.Name for f in FilteredElementCollector(doc).OfClass(Family))
missing = [n for n in fam_files if n not in loaded]

if missing:
    handler = FamilyLoaderOptionsHandler()
    t = Transaction(doc, 'Load Duct Size Family')
    t.Start()
    try:
        for n in missing:
            doc.LoadFamily(fam_files[n], handler)
        t.Commit()
    except Exception:
        t.RollBack()
        raise

# Place the first matching type only
symb = None
for s in (FilteredElementCollector(doc)
          .OfCategory(BuiltInCategory.OST_FabricationDuctworkTags)
          .OfClass(FamilySymbol)):
    if s.Family.Name == symbName:
        symb = s
        break

if symb and uidoc.CanPlaceElementType(symb):
    uidoc.PostRequestForElementTypePlacement(symb)