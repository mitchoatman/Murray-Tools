# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
import System

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    BuiltInCategory,
    Family,
    FamilyInstance,
    ViewType,
    XYZ,
    FabricationPart
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import os
import sys
import re


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

script_dir = os.path.dirname(os.path.abspath(__file__))
family_path = os.path.join(script_dir, 'Pipe Riser.rfa')

FAMILY_NAME = 'Pipe Riser'
FAMILY_TYPE = 'Pipe Riser'


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
# Fab pipe helpers
# --------------------------------------------------
def get_part_pattern_number(element):
    try:
        p = element.LookupParameter('Part Pattern Number')
        if p:
            return p.AsInteger()
    except:
        pass
    return None


def is_cid_2041_or_straight_pipe(element):
    try:
        if element.ItemCustomId == 2041:
            return True
    except:
        pass

    pattern = get_part_pattern_number(element)
    return pattern in (2041, 866, 40)


def is_vertical_fab_pipe(element, z_threshold=0.99):
    try:
        conns = list(element.ConnectorManager.Connectors)
        if len(conns) < 2:
            return False

        p1 = conns[0].Origin
        p2 = conns[1].Origin
        v = p2.Subtract(p1)

        if v.GetLength() < 0.0001:
            return False

        v = v.Normalize()
        return abs(v.Z) > z_threshold
    except:
        return False


def is_valid_pipe_riser(element):
    if not isinstance(element, FabricationPart):
        return False

    try:
        if not element.Category:
            return False
        if element.Category.Id.IntegerValue != int(BuiltInCategory.OST_FabricationPipework):
            return False
    except:
        return False

    if not is_cid_2041_or_straight_pipe(element):
        return False

    if not is_vertical_fab_pipe(element):
        return False

    return True


# --------------------------------------------------
# Selection filter
# --------------------------------------------------
class FabricationPipeRiserFilter(ISelectionFilter):
    def AllowElement(self, e):
        return is_valid_pipe_riser(e)

    def AllowReference(self, ref, point):
        return False


# --------------------------------------------------
# Family manager
# --------------------------------------------------
class PipeRiserFamilyManager(object):
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
# Helpers
# --------------------------------------------------
def fraction_to_float(text):
    if '/' in text:
        a, b = text.split('/')
        return float(a) / float(b)
    return float(text)


def parse_size_inches(size_str):
    if not size_str:
        return None

    s = size_str.strip().lower()
    s = s.replace('ø', '').replace('"', '').strip()

    m = re.search(r'(\d+\s+\d+/\d+|\d+/\d+|\d+(?:\.\d+)?)', s)
    if not m:
        return None

    token = m.group(1).strip()

    if ' ' in token and '/' in token:
        whole, frac = token.split()
        return float(whole) + fraction_to_float(frac)
    elif '/' in token:
        return fraction_to_float(token)
    else:
        return float(token)


def get_pipe_center(pipe):
    try:
        conns = list(pipe.ConnectorManager.Connectors)
        if conns:
            x = sum(c.Origin.X for c in conns) / len(conns)
            y = sum(c.Origin.Y for c in conns) / len(conns)
            z = sum(c.Origin.Z for c in conns) / len(conns)
            return XYZ(x, y, z)
    except:
        pass

    try:
        curve = pipe.Location.Curve
        return curve.Evaluate(0.5, True)
    except:
        pass

    try:
        bbox = pipe.get_BoundingBox(None)
        if bbox:
            return XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0
            )
    except:
        pass

    return None


def get_pipe_diameter_feet(pipe):
    try:
        conns = list(pipe.ConnectorManager.Connectors)
        for c in conns:
            try:
                if c.Radius > 0:
                    return c.Radius * 2.0
            except:
                pass
    except:
        pass

    size_str = None

    try:
        p = pipe.get_Parameter(DB.BuiltInParameter.RBS_REFERENCE_OVERALLSIZE)
        if p:
            size_str = p.AsString()
    except:
        pass

    if not size_str:
        try:
            p = pipe.LookupParameter("Overall Size")
            if p:
                size_str = p.AsString()
        except:
            pass

    inches = parse_size_inches(size_str)
    if inches is not None:
        return inches / 12.0

    try:
        p = pipe.LookupParameter("Diameter")
        if p:
            return p.AsDouble()
    except:
        pass

    return None


def get_target_level_from_view(view):
    try:
        return view.GenLevel
    except:
        return None


def get_target_z_from_view(view):
    try:
        if view.SketchPlane:
            return view.SketchPlane.GetPlane().Origin.Z
    except:
        pass

    try:
        if view.GenLevel:
            return view.GenLevel.ProjectElevation
    except:
        pass

    return None


def get_family_instance_point_xy(el):
    try:
        loc = getattr(el.Location, 'Point', None)
        if loc:
            return XYZ(loc.X, loc.Y, 0)
    except:
        pass

    try:
        bbox = el.get_BoundingBox(None)
        if bbox:
            cx = (bbox.Min.X + bbox.Max.X) / 2.0
            cy = (bbox.Min.Y + bbox.Max.Y) / 2.0
            return XYZ(cx, cy, 0)
    except:
        pass

    return None


def make_xy_key(point, tol=0.1):
    return (int(round(point.X / tol)), int(round(point.Y / tol)))


def get_existing_pipe_riser_instances(doc, family_name):
    existing = []

    for el in FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType():
        try:
            if el.Symbol and el.Symbol.Family and el.Symbol.Family.Name == family_name:
                existing.append(el)
        except:
            pass

    return existing


def build_existing_xy_map(existing_instances, tol=0.1):
    xy_map = {}

    for el in existing_instances:
        pt = get_family_instance_point_xy(el)
        if pt:
            xy_map[make_xy_key(pt, tol)] = el

    return xy_map


def set_diameter_if_possible(inst, dia_ft):
    if dia_ft is None:
        return False

    p = inst.LookupParameter("Diameter")
    if p and not p.IsReadOnly:
        try:
            p.Set(dia_ft)
            return True
        except:
            pass
    return False


# --------------------------------------------------
# Main
# --------------------------------------------------
def main():
    family_manager = PipeRiserFamilyManager(doc, FAMILY_NAME, family_path)
    famsymb = family_manager.get_ready_symbol(FAMILY_TYPE)

    if not famsymb:
        return

    try:
        picked = uidoc.Selection.PickObjects(
            ObjectType.Element,
            FabricationPipeRiserFilter(),
            "Select vertical CID 2041 / straight fabrication pipe risers"
        )
    except:
        return

    pipes = [doc.GetElement(r.ElementId) for r in picked]
    pipes = [p for p in pipes if p and is_valid_pipe_riser(p)]

    if not pipes:
        show_message("Error", "No valid vertical fabrication pipe risers were selected.")
        return

    target_level = get_target_level_from_view(view)
    target_z = get_target_z_from_view(view)

    if target_z is None:
        show_message("Error", "No valid work plane or level found in active view.")
        return

    xy_tol = 0.1

    existing_risers = get_existing_pipe_riser_instances(doc, FAMILY_NAME)
    existing_xy_map = build_existing_xy_map(existing_risers, xy_tol)

    t = Transaction(doc, "Place Pipe Risers")
    t.Start()

    placed = 0
    updated_existing = 0
    no_location = 0
    no_diameter = 0
    failed = 0

    for pipe in pipes:
        pipe_pt = get_pipe_center(pipe)
        dia_ft = get_pipe_diameter_feet(pipe)

        if pipe_pt is None:
            no_location += 1
            continue

        place_pt = XYZ(pipe_pt.X, pipe_pt.Y, target_z)
        xy_key = make_xy_key(place_pt, xy_tol)

        existing_inst = existing_xy_map.get(xy_key)

        if existing_inst:
            if set_diameter_if_possible(existing_inst, dia_ft):
                updated_existing += 1
            else:
                no_diameter += 1
            continue

        try:
            inst = None

            if target_level:
                try:
                    inst = doc.Create.NewFamilyInstance(
                        place_pt,
                        famsymb,
                        target_level,
                        DB.Structure.StructuralType.NonStructural
                    )
                except:
                    inst = None

            if inst is None:
                try:
                    inst = doc.Create.NewFamilyInstance(
                        place_pt,
                        famsymb,
                        DB.Structure.StructuralType.NonStructural
                    )
                except:
                    inst = None

            if inst is None:
                failed += 1
                continue

            if not set_diameter_if_possible(inst, dia_ft):
                no_diameter += 1

            existing_xy_map[xy_key] = inst
            placed += 1

        except:
            failed += 1

    t.Commit()

    msg = "Placed: {}".format(placed)
    if updated_existing:
        msg += "\nUpdated existing at same XY: {}".format(updated_existing)
    if no_location:
        msg += "\nNo pipe location: {}".format(no_location)
    if no_diameter:
        msg += "\nCould not set Diameter: {}".format(no_diameter)
    if failed:
        msg += "\nPlacement failed: {}".format(failed)

    show_message("Result", msg)


if __name__ == '__main__':
    main()