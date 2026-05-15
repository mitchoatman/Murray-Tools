# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
import System

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    BuiltInCategory,
    FamilySymbol,
    Family,
    ViewType
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
import os
import sys


# --------------------------------------------------
# Basic environment
# --------------------------------------------------
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
uiapp = __revit__
view = uidoc.ActiveView

if view.ViewType == ViewType.ThreeD:
    TaskDialog.Show("Error", "Cannot use in 3D view.")
    sys.exit()

path, filename = os.path.split(__file__)
family_path = os.path.join(path, 'Round Floor Sleeve.rfa')

FAMILY_NAME = 'Round Floor Sleeve'
FAMILY_TYPE = 'Round Floor Sleeve'


def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        pass


# --------------------------------------------------
# Family load options
# --------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Project
        overwriteParameterValues.Value = False
        return True


def shared_family_dialog_fallback(sender, args):
    try:
        if isinstance(args, TaskDialogShowingEventArgs):
            msg = (args.Message or "").lower()
            dialog_id = (args.DialogId or "").lower()

            if ("shared" in msg and "already exists" in msg and "project" in msg) \
               or ("shared" in dialog_id and "family" in dialog_id):
                args.OverrideResult(1003)
    except:
        pass


# --------------------------------------------------
# Family manager
# --------------------------------------------------
class RoundFloorSleeveFamilyManager(object):
    def __init__(self, document, family_name, family_path):
        self.doc = document
        self.family_name = family_name
        self.family_path = family_path
        self.family = None
        self.symbol_cache = {}

    def get_family_by_name(self):
        if self.family and self.family.IsValidObject:
            return self.family

        for fam in FilteredElementCollector(self.doc).OfClass(Family):
            if fam.Name == self.family_name:
                self.family = fam
                return fam

        self.family = None
        return None

    def get_symbol_name(self, symbol):
        try:
            if symbol.Name:
                return symbol.Name
        except:
            pass

        try:
            p = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if p:
                return p.AsString()
        except:
            pass

        return None

    def load_family_if_missing(self):
        fam = self.get_family_by_name()
        if fam:
            return fam

        if not os.path.exists(self.family_path):
            show_message("Error", "Family file not found:\n{}".format(self.family_path))
            return None

        t = None
        uiapp.DialogBoxShowing += shared_family_dialog_fallback
        try:
            t = Transaction(self.doc, "Load {} Family".format(self.family_name))
            t.Start()

            loaded_family_ref = clr.Reference[Family]()
            result = self.doc.LoadFamily(
                self.family_path,
                FamilyLoaderOptionsHandler(),
                loaded_family_ref
            )

            t.Commit()

            if result and loaded_family_ref.Value:
                self.family = loaded_family_ref.Value
                return self.family

            return self.get_family_by_name()

        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            show_message("Error", "Family load error:\n{}".format(str(e)))
            return None

        finally:
            uiapp.DialogBoxShowing -= shared_family_dialog_fallback
            if t:
                t.Dispose()

    def build_symbol_cache(self):
        self.symbol_cache = {}

        fam = self.get_family_by_name()
        if not fam:
            return

        for symbol_id in fam.GetFamilySymbolIds():
            sym = self.doc.GetElement(symbol_id)
            if sym:
                type_name = self.get_symbol_name(sym)
                if type_name:
                    self.symbol_cache[type_name.strip().upper()] = sym

    def get_symbol_by_type_name(self, type_name):
        if not type_name:
            return None

        if not self.symbol_cache:
            self.build_symbol_cache()

        return self.symbol_cache.get(type_name.strip().upper())

    def activate_symbol_if_needed(self, symbol):
        if not symbol:
            return False

        if symbol.IsActive:
            return True

        t = None
        try:
            t = Transaction(self.doc, "Activate {} Symbol".format(self.family_name))
            t.Start()
            symbol.Activate()
            self.doc.Regenerate()
            t.Commit()
            return True
        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            show_message("Error", "Family symbol activation error:\n{}".format(str(e)))
            return False
        finally:
            if t:
                t.Dispose()

    def get_ready_symbol(self, type_name):
        fam = self.load_family_if_missing()
        if not fam:
            return None

        self.build_symbol_cache()

        sym = self.get_symbol_by_type_name(type_name)
        if not sym:
            show_message(
                "Error",
                "Type '{}' not found in family '{}'.".format(type_name, self.family_name)
            )
            return None

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


# --------------------------------------------------
# Main
# --------------------------------------------------
def main():
    family_manager = RoundFloorSleeveFamilyManager(doc, FAMILY_NAME, family_path)
    famsymb = family_manager.get_ready_symbol(FAMILY_TYPE)

    if not famsymb:
        return

    try:
        uidoc.PostRequestForElementTypePlacement(famsymb)
    except Exception as e:
        show_message("Error", "Could not start placement:\n{}".format(str(e)))


if __name__ == '__main__':
    main()