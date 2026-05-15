# -*- coding: utf-8 -*-
import os
import sys
import clr

import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, IFamilyLoadOptions, Family
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
from pyrevit import forms


doc = __revit__.ActiveUIDocument.Document
uiapp = __revit__

COMPLETE_FOLDER_PATH = r"C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES"


class FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        # Keep existing type parameter values
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        # Always use the shared family already in the project
        source.Value = DB.FamilySource.Project
        overwriteParameterValues.Value = False
        return True


def shared_family_dialog_fallback(sender, args):
    """Fallback in case Revit still throws the shared family dialog."""
    try:
        if isinstance(args, TaskDialogShowingEventArgs):
            msg = (args.Message or "").lower()
            dialog_id = (args.DialogId or "").lower()

            if ("shared" in msg and "already exists" in msg and "project" in msg) \
               or ("shared" in dialog_id and "family" in dialog_id):
                # 3rd button = "Use the existing sub-component family that is in the project"
                args.OverrideResult(1003)
    except:
        pass


def get_family_by_name(document, family_name):
    collector = DB.FilteredElementCollector(document).OfClass(DB.Family)
    for fam in collector:
        if fam.Name == family_name:
            return fam
    return None


def load_family_robust(document, family_path):
    if not family_path:
        forms.alert("Operation canceled by user.", exitscript=True)

    if not os.path.exists(family_path):
        forms.alert("Family file not found:\n{}".format(family_path), exitscript=True)

    if not family_path.lower().endswith(".rfa"):
        forms.alert("Selected file is not an RFA:\n{}".format(family_path), exitscript=True)

    family_name = os.path.splitext(os.path.basename(family_path))[0]
    existing_family = get_family_by_name(document, family_name)

    t = None
    uiapp.DialogBoxShowing += shared_family_dialog_fallback
    try:
        t = Transaction(document, "Load Family")
        t.Start()

        load_options = FamilyLoadOptions()
        loaded_family_ref = clr.StrongBox[Family]()
        success = document.LoadFamily(family_path, load_options, loaded_family_ref)

        if success and loaded_family_ref.Value:
            t.Commit()
            return loaded_family_ref.Value, True

        # If LoadFamily returns False, the family may already exist unchanged
        t.Commit()

        fallback_family = get_family_by_name(document, family_name)
        if fallback_family:
            return fallback_family, False

        return None, False

    except Exception as ex:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise ex

    finally:
        uiapp.DialogBoxShowing -= shared_family_dialog_fallback
        if t:
            t.Dispose()


try:
    family_path = forms.pick_file(
        file_ext='rfa',
        init_dir=COMPLETE_FOLDER_PATH,
        multi_file=False,
        title='Select the family to insert'
    )

    family, newly_loaded = load_family_robust(doc, family_path)

    if family:
        if newly_loaded:
            forms.show_balloon("Load Family", "Loaded: {}".format(family.Name))
        else:
            forms.show_balloon("Load Family", "Using existing / already-loaded family: {}".format(family.Name))
    else:
        forms.alert("Failed to load family:\n{}".format(family_path), exitscript=True)

except Exception as e:
    forms.alert("An error occurred:\n{}".format(str(e)), exitscript=True)