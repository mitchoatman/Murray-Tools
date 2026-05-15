from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, Family
from Autodesk.Revit.UI import TaskDialog
import os
import sys

path, filename = os.path.split(__file__)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Change only this:
ACTIVE_TAG = "BOI"   # Options: "BOI", "BOP", "CL"

# Family/type map
TAG_CONFIG = {
    "BOI": {
        "family_name": "Fabrication Pipe - BOI Elevation and Pipe Size Tag",
        "type_name": "BOI",
        "file": os.path.join(path, "Fabrication Pipe - BOI Elevation and Pipe Size Tag.rfa")
    },
    "BOP": {
        "family_name": "Fabrication Pipe - BOP Elevation and Pipe Size Tag",
        "type_name": "BOP",
        "file": os.path.join(path, "Fabrication Pipe - BOP Elevation and Pipe Size Tag.rfa")
    },
    "CL": {
        "family_name": "Fabrication Pipe - CL Elevation and Pipe Size Tag",
        "type_name": "CL",
        "file": os.path.join(path, "Fabrication Pipe - CL Elevation and Pipe Size Tag.rfa")
    }
}


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


def load_family_if_missing(document, family_name, family_path):
    family = get_family_by_name(document, family_name)
    if family:
        return family

    if not os.path.exists(family_path):
        TaskDialog.Show("Error", "Family file not found:\n{}".format(family_path))
        sys.exit()

    t = Transaction(document, "Load Family - {}".format(family_name))
    t.Start()
    try:
        fload_handler = FamilyLoaderOptionsHandler()
        result = document.LoadFamily(family_path, fload_handler)
        if not result:
            t.RollBack()
            TaskDialog.Show("Error", "Failed to load family '{}'.".format(family_name))
            sys.exit()
        t.Commit()
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        TaskDialog.Show("Error", "Error loading family '{}':\n{}".format(family_name, str(e)))
        sys.exit()

    family = get_family_by_name(document, family_name)
    if not family:
        TaskDialog.Show("Error", "Family '{}' not found in project after loading.".format(family_name))
        sys.exit()

    return family


if ACTIVE_TAG not in TAG_CONFIG:
    TaskDialog.Show(
        "Error",
        "Invalid ACTIVE_TAG '{}'. Use one of: {}".format(
            ACTIVE_TAG, ", ".join(sorted(TAG_CONFIG.keys()))
        )
    )
    sys.exit()

# Load all families for future use
for key, config in TAG_CONFIG.items():
    load_family_if_missing(doc, config["family_name"], config["file"])

# Use only the selected one
active_config = TAG_CONFIG[ACTIVE_TAG]
family = get_family_by_name(doc, active_config["family_name"])

if not family:
    TaskDialog.Show("Error", "Family '{}' not found in project.".format(active_config["family_name"]))
    sys.exit()

target_symbol = get_symbol_by_type_name(doc, family, active_config["type_name"])

if not target_symbol:
    TaskDialog.Show(
        "Error",
        "Could not find type '{}' in family '{}'.".format(
            active_config["type_name"],
            active_config["family_name"]
        )
    )
    sys.exit()

if not target_symbol.IsActive:
    t = Transaction(doc, "Activate Family Symbol")
    t.Start()
    try:
        target_symbol.Activate()
        doc.Regenerate()
        t.Commit()
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        TaskDialog.Show("Error", "Error activating symbol:\n{}".format(str(e)))
        sys.exit()

uidoc.PostRequestForElementTypePlacement(target_symbol)