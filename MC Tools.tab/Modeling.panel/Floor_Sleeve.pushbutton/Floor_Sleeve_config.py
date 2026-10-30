# -*- coding: utf-8 -*-
"""Load (if needed) and place the Round Floor Sleeve family."""
import os

import clr
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    Family,
    ViewType
)
from Autodesk.Revit.UI import TaskDialog


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

FAMILY_NAME = 'Round Floor Sleeve'
FAMILY_TYPE = 'Round Floor Sleeve'
FAMILY_FILE = FAMILY_NAME + '.rfa'


def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except Exception:
        pass


def get_family_paths():
    """Candidate family files, in order: <script folder>/<Revit version>/ first,
    then next to the script. Keep one family per Revit version, saved in that
    version, so Revit never has to upgrade it during the load."""
    script_dir = os.path.dirname(__file__)
    version = doc.Application.VersionNumber
    return [
        os.path.join(script_dir, version, FAMILY_FILE),
        os.path.join(script_dir, FAMILY_FILE),
    ]


class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    """Keep the project's existing family and shared nested families on reload."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Project
        overwriteParameterValues.Value = False
        return True


def find_family():
    for fam in FilteredElementCollector(doc).OfClass(Family):
        if fam.Name == FAMILY_NAME:
            return fam
    return None


def load_family():
    paths = get_family_paths()
    family_path = next((p for p in paths if os.path.exists(p)), None)
    if not family_path:
        show_message(
            "Error",
            "Family file not found. Looked in:\n{}".format("\n".join(paths))
        )
        return None

    t = Transaction(doc, "Load {} Family".format(FAMILY_NAME))
    try:
        t.Start()
        family_ref = clr.Reference[Family]()
        doc.LoadFamily(family_path, FamilyLoaderOptionsHandler(), family_ref)
        t.Commit()
        return family_ref.Value or find_family()
    except Exception as e:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        show_message("Error", "Family load error:\n{}".format(e))
        return None
    finally:
        t.Dispose()


def find_symbol(family, type_name):
    target = type_name.strip().upper()
    for symbol_id in family.GetFamilySymbolIds():
        symbol = doc.GetElement(symbol_id)
        if not symbol:
            continue
        param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        name = param.AsString() if param else None
        if name and name.strip().upper() == target:
            return symbol
    return None


def activate_symbol(symbol):
    if symbol.IsActive:
        return True

    t = Transaction(doc, "Activate {} Symbol".format(FAMILY_NAME))
    try:
        t.Start()
        symbol.Activate()
        t.Commit()
        return True
    except Exception as e:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        show_message("Error", "Family symbol activation error:\n{}".format(e))
        return False
    finally:
        t.Dispose()


def main():
    if doc.ActiveView.ViewType == ViewType.ThreeD:
        show_message("Error", "Cannot use in 3D view.")
        return

    family = find_family() or load_family()
    if not family:
        return

    symbol = find_symbol(family, FAMILY_TYPE)
    if not symbol:
        show_message(
            "Error",
            "Type '{}' not found in family '{}'.".format(FAMILY_TYPE, FAMILY_NAME)
        )
        return

    if not activate_symbol(symbol):
        return

    # Starts Revit's native Place Component command with this type.
    try:
        uidoc.PostRequestForElementTypePlacement(symbol)
    except Exception as e:
        show_message("Error", "Could not start placement:\n{}".format(e))


if __name__ == '__main__':
    main()