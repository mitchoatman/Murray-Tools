from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, Family, ViewType
from Autodesk.Revit.UI import TaskDialog, Selection
import os
import sys

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = uidoc.ActiveView

if view.ViewType == ViewType.ThreeD:
    TaskDialog.Show("Error", "Script cannot run in 3D view.")
    sys.exit()

path, filename = os.path.split(__file__)
family_path = os.path.join(path, 'Flow Arrow.rfa')
FamilyName = 'Flow Arrow'
FamilyType = 'Flow Arrow'


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


def get_symbol_by_type_name(document, family, type_name):
    if not family:
        return None

    for symbol_id in family.GetFamilySymbolIds():
        symbol = document.GetElement(symbol_id)
        if symbol:
            p = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if p and p.AsString() == type_name:
                return symbol
    return None


family = get_family_by_name(doc, FamilyName)

if not family:
    if not os.path.exists(family_path):
        TaskDialog.Show("Error", "Family file not found:\n{}".format(family_path))
        sys.exit()

    t = Transaction(doc, 'Load Flow Arrow Family')
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

fam_symb = get_symbol_by_type_name(doc, family, FamilyType)
if not fam_symb:
    TaskDialog.Show("Error", "Type '{}' not found in family '{}'.".format(FamilyType, FamilyName))
    sys.exit()

if not fam_symb.IsActive:
    t = Transaction(doc, 'Activate Flow Arrow Symbol')
    t.Start()
    try:
        fam_symb.Activate()
        doc.Regenerate()
        t.Commit()
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        TaskDialog.Show("Error", "Error activating symbol:\n{}".format(str(e)))
        sys.exit()

try:
    selected_refs = uidoc.Selection.PickObjects(
        Selection.ObjectType.Element,
        "Select elements to place Flow Arrow"
    )
except Exception:
    TaskDialog.Show("Error", "Selection cancelled or failed.")
    sys.exit()

t = Transaction(doc, 'Place Flow Arrow Instances')
t.Start()
try:
    for ref in selected_refs:
        elem = doc.GetElement(ref)
        bbox = elem.get_BoundingBox(doc.ActiveView)
        if bbox:
            center = (bbox.Min + bbox.Max) / 2.0
            doc.Create.NewFamilyInstance(center, fam_symb, doc.ActiveView)
    t.Commit()
except Exception as e:
    if t.HasStarted():
        t.RollBack()
    TaskDialog.Show("Error", "Error placing Flow Arrow instances:\n{}".format(str(e)))