from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, ViewType
from Autodesk.Revit.UI import TaskDialog, Selection
import os
import sys

# Check active view type
view = __revit__.ActiveUIDocument.ActiveView
if view.ViewType == ViewType.ThreeD:
    TaskDialog.Show("Error", "Script cannot run in 3D view.")
    sys.exit()

path, filename = os.path.split(__file__)
NewFilename = '\Flow Arrow.rfa'

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
FamilyName = 'Flow Arrow'
Fam_is_in_project = any(f.Name == FamilyName for f in families)

family_pathCC = path + NewFilename

# Load family if not in project
t = Transaction(doc, 'Load Flow Arrow Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()

# Get family symbol
collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).OfClass(FamilySymbol)
fam_symb = None
for fs in collector:
    if fs.Family.Name == 'Flow Arrow':
        fam_symb = fs
        break

if fam_symb and not fam_symb.IsActive:
    fam_symb.Activate()

# Prompt user for element selection
try:
    selected_refs = uidoc.Selection.PickObjects(Selection.ObjectType.Element, "Select elements to place Flow Arrow")
except:
    TaskDialog.Show("Error", "Selection cancelled or failed.")
    sys.exit()

# Place family at center of each selected element's bounding box
t = Transaction(doc, 'Place Flow Arrow Instances')
t.Start()
for ref in selected_refs:
    elem = doc.GetElement(ref)
    bbox = elem.get_BoundingBox(doc.ActiveView)
    if bbox:
        center = (bbox.Min + bbox.Max) / 2
        doc.Create.NewFamilyInstance(center, fam_symb, doc.ActiveView)
t.Commit()