# -*- coding: utf-8 -*-

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    Transaction,
    Line,
    XYZ,
    ViewType
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs

from fractions import Fraction
import clr
import re
import os

from Parameters.Add_SharedParameters import Shared_Params
Shared_Params()

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString
)

# Revit document objects
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
curview = doc.ActiveView

# File path setup
path, filename = os.path.split(__file__)
family_path = os.path.join(path, 'Pipe Riser.rfa')

FAMILY_NAME = 'Pipe Riser'
FAMILY_TYPE_NAME = 'Pipe Riser'
PIPE_CATEGORY_INT = int(DB.BuiltInCategory.OST_FabricationPipework)
SLEEVE_CATEGORY = DB.BuiltInCategory.OST_PipeAccessory


# --------------------------------------------------
# Family load options / dialog fallback
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
class PipeRiserFamilyManager(object):
    def __init__(self, document, target_family_name, target_family_path):
        self.doc = document
        self.family_name = target_family_name
        self.family_path = target_family_path
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
            TaskDialog.Show("Error", "Family file not found:\n{}".format(self.family_path))
            return None

        t = None
        uiapp.DialogBoxShowing += shared_family_dialog_fallback
        try:
            t = Transaction(self.doc, "Load {}".format(self.family_name))
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
            TaskDialog.Show("Error", "Error loading family '{}': {}".format(self.family_name, str(e)))
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
            t = Transaction(self.doc, "Activate {}".format(self.get_symbol_name(symbol) or "Family Symbol"))
            t.Start()
            symbol.Activate()
            self.doc.Regenerate()
            t.Commit()
            return True

        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            TaskDialog.Show("Error", "Error activating family symbol: {}".format(str(e)))
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
            TaskDialog.Show(
                "Error",
                "Type '{}' not found in family '{}'.".format(type_name, self.family_name)
            )
            return None

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


def is_fabrication_pipe(elem):
    try:
        return elem and elem.Category and elem.Category.Id.IntegerValue == PIPE_CATEGORY_INT
    except:
        return False


def get_pipe_connectors(pipe):
    try:
        return list(pipe.ConnectorManager.Connectors)
    except:
        return []


def get_pipe_endpoints(pipe):
    try:
        loc = pipe.Location
        if loc and hasattr(loc, "Curve") and loc.Curve:
            return loc.Curve.GetEndPoint(0), loc.Curve.GetEndPoint(1)
    except:
        pass

    connectors = get_pipe_connectors(pipe)
    if len(connectors) >= 2:
        pts = [c.Origin for c in connectors]
        max_pair = None
        max_dist = -1.0
        for i in range(len(pts)):
            for j in range(i + 1, len(pts)):
                d = pts[i].DistanceTo(pts[j])
                if d > max_dist:
                    max_dist = d
                    max_pair = (pts[i], pts[j])
        if max_pair:
            return max_pair

    return None, None


def get_pipe_centerline(pipe):
    p0, p1 = get_pipe_endpoints(pipe)
    if not p0 or not p1 or p0.IsAlmostEqualTo(p1):
        return None
    return Line.CreateBound(p0, p1)


def is_vertical_pipe(pipe, z_tolerance=0.99):
    p0, p1 = get_pipe_endpoints(pipe)
    if not p0 or not p1:
        return False

    vec = p1 - p0
    if vec.GetLength() < 1e-9:
        return False

    direction = vec.Normalize()
    return abs(direction.Z) >= z_tolerance


def get_pipe_intersections(pipe, level, z_tolerance=1e-6):
    centerline = get_pipe_centerline(pipe)
    if not centerline:
        return []

    p0 = centerline.GetEndPoint(0)
    p1 = centerline.GetEndPoint(1)

    z0 = p0.Z
    z1 = p1.Z
    zl = level.Elevation

    if abs(z1 - z0) < z_tolerance:
        return []

    min_z = min(z0, z1)
    max_z = max(z0, z1)

    if zl < min_z - z_tolerance or zl > max_z + z_tolerance:
        return []

    ratio = (zl - z0) / (z1 - z0)
    x = p0.X + (p1.X - p0.X) * ratio
    y = p0.Y + (p1.Y - p0.Y) * ratio
    pt = XYZ(x, y, zl)

    return [pt]


def get_world_bbox_min_max(bbox):
    if not bbox:
        return None, None

    t = bbox.Transform
    min_pt = bbox.Min
    max_pt = bbox.Max

    corners = [
        XYZ(min_pt.X, min_pt.Y, min_pt.Z),
        XYZ(min_pt.X, min_pt.Y, max_pt.Z),
        XYZ(min_pt.X, max_pt.Y, min_pt.Z),
        XYZ(min_pt.X, max_pt.Y, max_pt.Z),
        XYZ(max_pt.X, min_pt.Y, min_pt.Z),
        XYZ(max_pt.X, min_pt.Y, max_pt.Z),
        XYZ(max_pt.X, max_pt.Y, min_pt.Z),
        XYZ(max_pt.X, max_pt.Y, max_pt.Z),
    ]

    world = [t.OfPoint(c) for c in corners]

    xs = [p.X for p in world]
    ys = [p.Y for p in world]
    zs = [p.Z for p in world]

    return XYZ(min(xs), min(ys), min(zs)), XYZ(max(xs), max(ys), max(zs))


def point_in_bbox(point, min_pt, max_pt, tol=1e-6):
    if not min_pt or not max_pt:
        return True

    return (
        min_pt.X - tol <= point.X <= max_pt.X + tol and
        min_pt.Y - tol <= point.Y <= max_pt.Y + tol and
        min_pt.Z - tol <= point.Z <= max_pt.Z + tol
    )


def get_existing_sleeves(document):
    sleeves = []
    for elem in FilteredElementCollector(document).OfCategory(SLEEVE_CATEGORY).WhereElementIsNotElementType():
        try:
            if elem.Symbol and elem.Symbol.Family and elem.Symbol.Family.Name == FAMILY_NAME:
                sleeves.append(elem)
        except:
            pass
    return sleeves


def is_duplicate_sleeve(intersection_point, existing_sleeves, tolerance=0.001):
    for sleeve in existing_sleeves:
        try:
            loc = sleeve.Location
            if not loc or not hasattr(loc, "Point") or not loc.Point:
                continue
            sleeve_point = loc.Point
            if (abs(sleeve_point.X - intersection_point.X) < tolerance and
                abs(sleeve_point.Y - intersection_point.Y) < tolerance and
                abs(sleeve_point.Z - intersection_point.Z) < tolerance):
                return True
        except:
            pass
    return False


def parse_pipe_diameter_feet(pipe):
    raw = None

    try:
        p = pipe.get_Parameter(DB.BuiltInParameter.RBS_REFERENCE_OVERALLSIZE)
        if p:
            raw = p.AsString()
    except:
        pass

    if not raw:
        try:
            raw = get_parameter_value_by_name_AsString(pipe, "Overall Size")
        except:
            raw = None

    if not raw:
        return 0.5 / 12.0

    cleaned = raw.strip().replace('"', '')

    try:
        return float(cleaned) / 12.0
    except:
        pass

    mixed_match = re.match(r'^\s*(\d+)\s*[- ]\s*(\d+/\d+)\s*$', cleaned)
    frac_match = re.match(r'^\s*(\d+/\d+)\s*$', cleaned)

    try:
        if mixed_match:
            whole, frac = mixed_match.groups()
            inches = float(whole) + float(Fraction(frac))
            return inches / 12.0
        elif frac_match:
            inches = float(Fraction(frac_match.group(1)))
            return inches / 12.0
    except:
        pass

    return 0.5 / 12.0


def safe_set_param(instance, param_name, value):
    try:
        p = instance.LookupParameter(param_name)
        if not p or p.IsReadOnly:
            return False

        if p.StorageType == DB.StorageType.String:
            p.Set("" if value is None else str(value))
            return True
        elif p.StorageType == DB.StorageType.Double and isinstance(value, (int, float)):
            p.Set(float(value))
            return True
        elif p.StorageType == DB.StorageType.ElementId and isinstance(value, DB.ElementId):
            p.Set(value)
            return True
    except:
        pass
    return False


def copy_pipe_values_to_sleeve(pipe, sleeve_instance):
    mapping = {
        'FP_Product Entry': 'Overall Size',
        'FP_Service Name': 'Fabrication Service Name',
        'FP_Service Abbreviation': 'Fabrication Service Abbreviation'
    }

    warnings = []

    for sleeve_param, pipe_param in mapping.items():
        try:
            value = get_parameter_value_by_name_AsString(pipe, pipe_param)
            if value is None:
                value = ""
                warnings.append("Missing pipe parameter '{}'".format(pipe_param))
            set_parameter_by_name(sleeve_instance, sleeve_param, value)
        except:
            try:
                safe_set_param(sleeve_instance, sleeve_param, "")
            except:
                pass
            warnings.append("Failed mapping '{}' -> '{}'".format(pipe_param, sleeve_param))

    return warnings


def place_sleeve_at_intersection(pipe, intersection_point, family_symbol, level, existing_sleeves):
    if is_duplicate_sleeve(intersection_point, existing_sleeves):
        return None, []

    new_instance = doc.Create.NewFamilyInstance(
        intersection_point,
        family_symbol,
        level,
        DB.Structure.StructuralType.NonStructural
    )

    warnings = []

    diameter = parse_pipe_diameter_feet(pipe)
    if not safe_set_param(new_instance, 'Diameter', diameter):
        warnings.append("Could not set 'Diameter'")

    if not safe_set_param(new_instance, 'Schedule Level', level.Id):
        warnings.append("Could not set 'Schedule Level'")

    warnings.extend(copy_pipe_values_to_sleeve(pipe, new_instance))

    existing_sleeves.append(new_instance)
    return new_instance, warnings


def get_target_pipes():
    selected_ids = uidoc.Selection.GetElementIds()

    if selected_ids.Count > 0:
        pipes = []
        for eid in selected_ids:
            elem = doc.GetElement(eid)
            if is_fabrication_pipe(elem):
                pipes.append(elem)
        return pipes, True

    pipes = FilteredElementCollector(doc, curview.Id) \
        .OfCategory(DB.BuiltInCategory.OST_FabricationPipework) \
        .WhereElementIsNotElementType() \
        .ToElements()

    return list(pipes), False


def get_visible_levels_for_view(view, levels):
    if view.ViewType != ViewType.ThreeD:
        current_level = view.GenLevel
        return [current_level] if current_level else []

    if view.IsSectionBoxActive:
        bbox = view.GetSectionBox()
        min_pt, max_pt = get_world_bbox_min_max(bbox)
        visible_levels = []
        for lvl in levels:
            z = lvl.Elevation
            if min_pt.Z <= z <= max_pt.Z:
                visible_levels.append(lvl)
        return visible_levels

    return list(levels)


def main():
    if curview.ViewType not in [ViewType.ThreeD, ViewType.FloorPlan]:
        TaskDialog.Show("Error", "Run this script from a 3D view or Floor Plan view.")
        return

    family_manager = PipeRiserFamilyManager(doc, FAMILY_NAME, family_path)
    family_symbol = family_manager.get_ready_symbol(FAMILY_TYPE_NAME)
    if not family_symbol:
        TaskDialog.Show("Error", "Failed to load family symbol.\n{}".format(family_path))
        return

    all_levels = list(FilteredElementCollector(doc).OfClass(DB.Level).ToElements())
    existing_sleeves = get_existing_sleeves(doc)

    pipes, using_selection = get_target_pipes()
    if not pipes:
        TaskDialog.Show("Info", "No fabrication pipes found.")
        return

    visible_levels = get_visible_levels_for_view(curview, all_levels)
    if not visible_levels:
        TaskDialog.Show("Error", "No valid level(s) found for the active view.")
        return

    section_min = None
    section_max = None
    if curview.ViewType == ViewType.ThreeD and curview.IsSectionBoxActive:
        section_min, section_max = get_world_bbox_min_max(curview.GetSectionBox())

    placed_count = 0
    skipped_non_vertical = 0
    warnings_all = []

    t = Transaction(doc, 'Place Pipe Risers at Intersections')
    t.Start()
    try:
        if curview.ViewType == ViewType.FloorPlan:
            current_level = curview.GenLevel
            if not current_level:
                TaskDialog.Show("Error", "No level associated with the current floor plan view.")
                t.RollBack()
                return

            for pipe in pipes:
                if not is_vertical_pipe(pipe):
                    skipped_non_vertical += 1
                    continue

                for point in get_pipe_intersections(pipe, current_level):
                    new_sleeve, warns = place_sleeve_at_intersection(
                        pipe, point, family_symbol, current_level, existing_sleeves
                    )
                    if new_sleeve:
                        placed_count += 1
                    warnings_all.extend(warns)

        else:
            visible_pipe_ids = set([p.Id for p in FilteredElementCollector(doc, curview.Id)
                                    .OfCategory(DB.BuiltInCategory.OST_FabricationPipework)
                                    .WhereElementIsNotElementType()
                                    .ToElements()])

            for pipe in pipes:
                if using_selection and pipe.Id not in visible_pipe_ids:
                    continue

                if not is_vertical_pipe(pipe):
                    skipped_non_vertical += 1
                    continue

                for level in visible_levels:
                    for point in get_pipe_intersections(pipe, level):
                        if not point_in_bbox(point, section_min, section_max):
                            continue

                        new_sleeve, warns = place_sleeve_at_intersection(
                            pipe, point, family_symbol, level, existing_sleeves
                        )
                        if new_sleeve:
                            placed_count += 1
                        warnings_all.extend(warns)

        t.Commit()

    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        TaskDialog.Show("Error", "Error placing pipe risers: {}".format(str(e)))
        return

    summary = []
    summary.append("Placed: {}".format(placed_count))
    if skipped_non_vertical:
        summary.append("Skipped non-vertical pipes: {}".format(skipped_non_vertical))

    unique_warnings = sorted(set(warnings_all))
    if unique_warnings:
        summary.append("")
        summary.append("Warnings:")
        summary.extend(unique_warnings[:10])
        if len(unique_warnings) > 10:
            summary.append("...and {} more".format(len(unique_warnings) - 10))

    TaskDialog.Show("Pipe Riser", "\n".join(summary))


if __name__ == '__main__':
    main()