# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
import System

import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import (
    IFamilyLoadOptions,
    FamilySource,
    Transaction,
    FilteredElementCollector,
    Family,
    TransactionGroup,
    BuiltInCategory,
    FamilySymbol,
    BuiltInParameter,
    Reference,
    IndependentTag,
    TagMode,
    TagOrientation,
    ViewType,
    View3D,
    ElementId
)
import os


# --------------------------------------------------
# Basic environment
# --------------------------------------------------
path, filename = os.path.split(__file__)
family_pathCC = os.path.join(path, 'Fabrication Hanger - Pointload Tag.rfa')

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

FamilyName = 'Fabrication Hanger - Pointload Tag'
FamilyType = 'CIRCLE'


def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        pass


# --------------------------------------------------
# View validation
# --------------------------------------------------
if curview.ViewType == ViewType.ThreeD:
    view3d = curview
    if not view3d.IsLocked:
        show_message(
            "Error",
            "This script cannot be run in an unlocked 3D view. Please lock the 3D view or switch to a plan, section, or elevation view and try again."
        )
        raise SystemExit("Script terminated: Unlocked 3D view detected.")


# --------------------------------------------------
# Robust family loading
# --------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        # Use the family file so missing/new types in the RFA get brought in
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


def shared_family_dialog_fallback(sender, args):
    try:
        if isinstance(args, TaskDialogShowingEventArgs):
            msg = (args.Message or "").lower()
            dialog_id = (args.DialogId or "").lower()

            if ("shared" in msg and "already exists" in msg) \
               or ("shared" in dialog_id and "family" in dialog_id):
                # Prefer family version
                args.OverrideResult(1001)
    except:
        pass


class PointloadTagFamilyManager(object):
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

    def load_family_from_disk(self, transaction_name):
        if not os.path.exists(self.family_path):
            show_message("Error", "Family file not found at path:\n{}".format(self.family_path))
            return None

        t = None
        uiapp.DialogBoxShowing += shared_family_dialog_fallback
        try:
            t = Transaction(self.doc, transaction_name)
            t.Start()

            loaded_family_ref = clr.Reference[Family]()
            result = self.doc.LoadFamily(
                self.family_path,
                FamilyLoaderOptionsHandler(),
                loaded_family_ref
            )

            self.doc.Regenerate()
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

    def ensure_family_and_type(self, type_name):
        fam = self.get_family_by_name()

        if not fam:
            fam = self.load_family_from_disk("Load {} Family".format(self.family_name))
            if not fam:
                return None

        self.build_symbol_cache()
        sym = self.get_symbol_by_type_name(type_name)
        if sym:
            return sym

        # Family exists but requested type is missing -> reload from disk
        fam = self.load_family_from_disk("Reload {} Family".format(self.family_name))
        if not fam:
            return None

        self.build_symbol_cache()
        sym = self.get_symbol_by_type_name(type_name)
        if not sym:
            show_message(
                "Error",
                "Type '{}' not found in family '{}' after reload.".format(type_name, self.family_name)
            )
            return None

        return sym

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
        sym = self.ensure_family_and_type(type_name)
        if not sym:
            return None

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


# --------------------------------------------------
# Existing tags
# --------------------------------------------------
def get_existing_tags(element):
    tags = FilteredElementCollector(doc, curview.Id) \
        .OfCategory(BuiltInCategory.OST_FabricationHangerTags) \
        .WhereElementIsNotElementType() \
        .ToElements()

    tagged_elements = set()

    try:
        if RevitINT >= 2022:
            for tag in tags:
                tagged_ids = tag.GetTaggedLocalElementIds()
                if tagged_ids and element.Id in tagged_ids:
                    tagged_elements.add(tag.Id)
        else:
            for tag in tags:
                if tag.TaggedLocalElementId == element.Id:
                    tagged_elements.add(tag.Id)
    except Exception as e:
        show_message("Warning", "Error retrieving existing tags: {}. Continuing with empty tag set.".format(str(e)))
        return tagged_elements

    return tagged_elements


def needs_tagging(element, rod_count, existing_tags):
    return len(existing_tags) < rod_count


# --------------------------------------------------
# Main
# --------------------------------------------------
def main():
    tg = None
    t = None

    try:
        family_manager = PointloadTagFamilyManager(doc, FamilyName, family_pathCC)
        famsymb = family_manager.get_ready_symbol(FamilyType)
        if not famsymb:
            raise SystemExit("Required family/type not available.")

        tg = TransactionGroup(doc, "Add Pointload Tags")
        tg.Start()

        hanger_collector = FilteredElementCollector(doc, curview.Id) \
            .OfCategory(BuiltInCategory.OST_FabricationHangers) \
            .WhereElementIsNotElementType()

        try:
            tag_cat_id = ElementId(BuiltInCategory.OST_FabricationHangerTags)
            if curview.GetCategoryHidden(tag_cat_id):
                show_message("Warning", "MEP Fabrication Hanger Tags category is not visible.")
        except:
            pass

        t = Transaction(doc, 'Tag Pointloads')
        t.Start()

        for hanger in hanger_collector:
            try:
                ref = Reference(hanger)
                rod_count = hanger.GetRodInfo().RodCount
                existing_tags = get_existing_tags(hanger)

                if needs_tagging(hanger, rod_count, existing_tags):
                    rod_info = hanger.GetRodInfo()
                    remaining_tags = rod_count - len(existing_tags)

                    for n in range(remaining_tags):
                        try:
                            rodloc = rod_info.GetRodEndPosition(n)
                            new_tag = IndependentTag.Create(
                                doc,
                                curview.Id,
                                ref,
                                False,
                                TagMode.TM_ADDBY_CATEGORY,
                                TagOrientation.Horizontal,
                                rodloc
                            )

                            if new_tag and new_tag.GetTypeId() != famsymb.Id:
                                try:
                                    new_tag.ChangeTypeId(famsymb.Id)
                                except:
                                    pass

                        except Exception as tag_error:
                            show_message(
                                "Error",
                                "Failed to create tag for element ID {}: {}. Skipping this tag.".format(
                                    hanger.Id, str(tag_error)
                                )
                            )
                            continue

            except Exception as hanger_error:
                show_message(
                    "Error",
                    "Error processing element ID {}: {}. Skipping this element.".format(
                        hanger.Id, str(hanger_error)
                    )
                )
                continue

        t.Commit()
        tg.Assimilate()

    except Exception as e:
        show_message(
            "Error",
            "An unexpected error occurred: {}. Please contact support or check the script.".format(str(e))
        )

        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if tg and tg.HasStarted() and not tg.HasEnded():
            tg.RollBack()

        raise SystemExit("Script terminated due to unexpected error: {}".format(str(e)))


if __name__ == '__main__':
    main()