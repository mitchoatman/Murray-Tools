from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, Family, ViewType
from Autodesk.Revit.UI import TaskDialog
import os
import sys

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = uidoc.ActiveView

if view.ViewType == ViewType.ThreeD:
    TaskDialog.Show("Error", "Cannot use in 3D view.")
    sys.exit()

path, filename = os.path.split(__file__)
family_path = os.path.join(path, 'North Arrow.rfa')

FamilyName = 'North Arrow'
FamilyType = 'North Arrow'


class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


def get_family_by_name(document, family_name):
    for fam in FilteredElementCollector(document).OfClass(Family):
        if fam.Name == family_name:
            return fam
    return None


family = get_family_by_name(doc, FamilyName)

if not family:
    if not os.path.exists(family_path):
        TaskDialog.Show("Error", "Family file not found:\n{}".format(family_path))
        sys.exit()

    t = Transaction(doc, 'Load North Arrow Family')
    t.Start()
    try:
        fload_handler = FamilyLoaderOptionsHandler()
        result = doc.LoadFamily(family_path, fload_handler)
        if not result:
            t.RollBack()
            TaskDialog.Show("Error", "Failed to load family '{}'.".format(FamilyName))
            sys.exit()
        t.Commit()
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        TaskDialog.Show("Error", "Error loading family:\n{}".format(str(e)))
        sys.exit()

    family = get_family_by_name(doc, FamilyName)

if not family:
    TaskDialog.Show("Error", "Family '{}' not found in project.".format(FamilyName))
    sys.exit()

target_symbol = None
for symbol_id in family.GetFamilySymbolIds():
    symbol = doc.GetElement(symbol_id)
    if symbol:
        p = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.AsString() == FamilyType:
            target_symbol = symbol
            break

if not target_symbol:
    TaskDialog.Show("Error", "Type '{}' not found in family '{}'.".format(FamilyType, FamilyName))
    sys.exit()

uidoc.PostRequestForElementTypePlacement(target_symbol)