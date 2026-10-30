# -*- coding: utf-8 -*-
import clr
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (FilteredElementCollector, Transaction,
                               BuiltInCategory, FamilySymbol, Family,
                               BasicFileInfo)
from Autodesk.Revit.UI import TaskDialog
import os
import tempfile
import datetime

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__

path, filename = os.path.split(__file__)
family_path = os.path.join(path, 'Fabrication Duct - Alias Tag.rfa')

FamilyName = 'Fabrication Duct - Alias Tag'
symbName = 'Fabrication Duct - Alias Tag'

LOG_PATH = os.path.join(tempfile.gettempdir(), 'alias_tag_debug.log')


def log(msg):
    try:
        f = open(LOG_PATH, 'a')
        f.write('{} {}\n'.format(datetime.datetime.now().strftime('%H:%M:%S'), msg))
        f.close()
    except:
        pass


def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        pass


class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


class StrictFailurePreprocessor(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        if failuresAccessor.GetSeverity() >= DB.FailureSeverity.Error:
            return DB.FailureProcessingResult.ProceedWithRollBack
        failuresAccessor.DeleteAllWarnings()
        return DB.FailureProcessingResult.Continue


class DialogGuard(object):
    def __init__(self):
        self.seen = []

    def handler(self, sender, args):
        try:
            self.seen.append(str(args.DialogId))
            args.OverrideResult(2)  # IDCANCEL, never leave a dialog open
        except:
            pass


def family_in_project():
    for f in FilteredElementCollector(doc).OfClass(Family):
        if f.Name == FamilyName:
            return True
    return False


def load_family():
    """Returns True if the family was loaded and committed."""
    if not os.path.exists(family_path):
        log("family file not found")
        show_message("Error", "Family file not found:\n{}".format(family_path))
        return False

    # Preflight: fail before touching Revit's load pipeline
    try:
        info = BasicFileInfo.Extract(family_path)
        if info.IsSavedInLaterVersion:
            log("family saved in later version: {}".format(info.Format))
            show_message("Error", "Family was saved in a newer Revit version ({}).".format(info.Format))
            return False
    except Exception as e:
        log("preflight failed: {}".format(e))
        show_message("Error", "Cannot read family file:\n{}".format(e))
        return False

    guard = DialogGuard()
    t = None
    uiapp.DialogBoxShowing += guard.handler
    try:
        t = Transaction(doc, 'Load Duct Alias Family')
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(StrictFailurePreprocessor())
        opts.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(opts)
        t.Start()

        log("LoadFamily start")
        ref = clr.Reference[Family]()
        ok = doc.LoadFamily(family_path, FamilyLoaderOptionsHandler(), ref)
        log("LoadFamily returned: {}, dialogs: {}".format(ok, guard.seen))

        if not ok or guard.seen:
            t.RollBack()
            show_message("Error", "Family load aborted. Dialogs raised: {}".format(
                ", ".join(guard.seen) or "none"))
            return False

        status = t.Commit()
        log("load commit status: {}".format(status))
        if status != DB.TransactionStatus.Committed:
            show_message("Error", "Family load did not commit.")
            return False
        return True

    except Exception as e:
        log("load exception: {}".format(e))
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        show_message("Error", "Family load error:\n{}".format(e))
        return False

    finally:
        uiapp.DialogBoxShowing -= guard.handler
        if t:
            t.Dispose()


def find_symbol():
    for s in (FilteredElementCollector(doc)
              .OfCategory(BuiltInCategory.OST_FabricationDuctworkTags)
              .OfClass(FamilySymbol)):
        if s.Family.Name == symbName:
            return s
    return None


def activate_symbol(symb):
    """Returns True if activation committed."""
    t = None
    try:
        log("activate start")
        t = Transaction(doc, 'Activate Alias Tag')
        t.Start()
        symb.Activate()
        doc.Regenerate()
        t.Commit()
        log("activate committed")
        return True
    except Exception as e:
        log("activate exception: {}".format(e))
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        show_message("Error", "Family symbol activation error:\n{}".format(e))
        return False
    finally:
        if t:
            t.Dispose()


def main():
    log("---- run start, view type: {}".format(uidoc.ActiveView.ViewType))
    changed = False

    if not family_in_project():
        if not load_family():
            return
        changed = True
    else:
        log("family already in project")

    symb = find_symbol()
    if not symb:
        log("symbol not found")
        show_message("Error", "Type '{}' not found in the Fabrication Duct Tags category.".format(symbName))
        return

    if not symb.IsActive:
        if not activate_symbol(symb):
            return
        changed = True

    # Two-phase: never post placement in the same run that changed the model
    if changed:
        log("model changed this run, placement deferred")
        show_message("Duct Alias Tag", "Family is loaded. Click the command again to place it.")
        return

    try:
        can_place = uidoc.CanPlaceElementType(symb)
        log("CanPlaceElementType: {}".format(can_place))
        if not can_place:
            show_message("Duct Alias Tag",
                         "Revit cannot place this tag type directly. Use Tag by Category instead.")
            return

        log("before post")
        uidoc.PostRequestForElementTypePlacement(symb)
        log("after post")
    except Exception as e:
        log("post exception: {}".format(e))
        show_message("Error", "Could not start placement:\n{}".format(e))


if __name__ == '__main__':
    main()