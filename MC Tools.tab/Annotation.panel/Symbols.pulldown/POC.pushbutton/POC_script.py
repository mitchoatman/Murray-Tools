from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, Family, ViewType
from Autodesk.Revit.UI import TaskDialog
import os
import sys

# --------------------------------------------------
# Basic environment
# --------------------------------------------------
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = uidoc.ActiveView

if view.ViewType == ViewType.ThreeD:
    TaskDialog.Show("Error", "Cannot use in 3D view.")
    sys.exit()

path, filename = os.path.split(__file__)
family_filename = 'POC Symbol.rfa'
family_path = os.path.join(path, family_filename)

FamilyName = 'POC Symbol'
FamilyType = 'POC Symbol'


# --------------------------------------------------
# Family load options
# --------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


# --------------------------------------------------
# Helpers
# --------------------------------------------------
def get_family_by_name(document, family_name):
    collector = FilteredElementCollector(document).OfClass(Family)
    for fam in collector:
        if fam.Name == family_name:
            return fam
    return None


def load_family_if_missing(document, family_name, family_path):
    fam = get_family_by_name(document, family_name)
    if fam:
        return fam

    if not os.path.exists(family_path):
        TaskDialog.Show("Error", "Family file not found:\n{}".format(family_path))
        return None

    t = Transaction(document, 'Load POC Symbol Family')
    t.Start()
    try:
        fload_handler = FamilyLoaderOptionsHandler()
        result = document.LoadFamily(family_path, fload_handler)

        if not result:
            t.RollBack()
            TaskDialog.Show("Error", "Failed to load family '{}'.".format(family_name))
            return None

        t.Commit()
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        TaskDialog.Show("Error", "Error loading family:\n{}".format(str(e)))
        return None

    return get_family_by_name(document, family_name)


def get_family_symbol_by_type_name(document, family, type_name):
    if not family:
        return None

    for symbol_id in family.GetFamilySymbolIds():
        symbol = document.GetElement(symbol_id)
        if symbol:
            p = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if p and p.AsString() == type_name:
                return symbol

    return None


# --------------------------------------------------
# Main
# --------------------------------------------------
family = load_family_if_missing(doc, FamilyName, family_path)
if not family:
    sys.exit()

symbol = get_family_symbol_by_type_name(doc, family, FamilyType)
if not symbol:
    TaskDialog.Show(
        "Error",
        "Type '{}' not found in family '{}'.".format(FamilyType, FamilyName)
    )
    sys.exit()

uidoc.PostRequestForElementTypePlacement(symbol)