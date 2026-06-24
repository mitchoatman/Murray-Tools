# -*- coding: utf-8 -*-
import clr
import os
import math

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    BuiltInCategory,
    Transaction,
    XYZ,
    ViewType,
    Transform,
    ProjectLocation,
    LocationCurve,
    Line,
    ElementTransformUtils,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs

from Parameters.Add_SharedParameters import Shared_Params
Shared_Params()

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString
)

# --------------------------------------------------
# Coordinate conversion helper
# --------------------------------------------------
class PointConverter(object):
    """Convert coordinates between internal / project / survey systems."""
    def __init__(self, x, y, z, coord_sys='internal', doc=None):
        if doc is None:
            doc = __revit__.ActiveUIDocument.Document
        self.doc = doc
        pt = XYZ(x, y, z)

        srv_trans = self._get_survey_transform()
        proj_trans = self._get_project_transform()

        if coord_sys.lower() == 'internal':
            self.internal = pt
            self.survey = srv_trans.Inverse.OfPoint(pt)
            self.project = proj_trans.Inverse.OfPoint(pt)
        elif coord_sys.lower() == 'project':
            self.project = pt
            self.internal = proj_trans.OfPoint(pt)
            self.survey = srv_trans.Inverse.OfPoint(self.internal)
        elif coord_sys.lower() == 'survey':
            self.survey = pt
            self.internal = srv_trans.OfPoint(pt)
            self.project = proj_trans.Inverse.OfPoint(self.internal)
        else:
            raise ValueError("coord_sys must be 'internal', 'project' or 'survey'")

    def _get_survey_transform(self):
        return self.doc.ActiveProjectLocation.GetTotalTransform()

    def _get_project_transform(self):
        collector = FilteredElementCollector(self.doc).OfClass(ProjectLocation).WhereElementIsNotElementType()
        for loc in collector:
            if loc.Name == "Project":
                return loc.GetTotalTransform()
        return Transform.Identity


# --------------------------------------------------
# Basic environment
# --------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
curview = doc.ActiveView

path, _ = os.path.split(__file__)

round_family_path = os.path.join(path, 'Round Floor Sleeve.rfa')
rect_family_path = os.path.join(path, 'Rectangular Floor Sleeve.rfa')

ROUND_FAMILY_NAME = 'Round Floor Sleeve'
RECT_FAMILY_NAME = 'Rectangular Floor Sleeve'

ROUND_TYPE_NAME = 'Round Floor Sleeve'
RECT_TYPE_NAME = 'Rectangular Sleeve'   # fallback to first type if this doesn't match

RECTANGULAR_ROTATION_OFFSET_DEGREES = 0.0


# --------------------------------------------------
# Helpers
# --------------------------------------------------
def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        pass


def safe_get_double_param_by_name(element, param_name, default=0.0):
    try:
        if element is None or not param_name:
            return default

        param = element.LookupParameter(param_name)
        if not param:
            return default

        if param.StorageType == DB.StorageType.Double:
            return param.AsDouble()

        if param.StorageType == DB.StorageType.Integer:
            return float(param.AsInteger())

        if param.StorageType == DB.StorageType.String:
            val = param.AsString()
            if val:
                return float(val)
    except:
        pass

    return default


def get_insulation_thickness(duct):
    possible_names = [
        'Insulation Thickness',
        'Insulation',
        'Ins Thickness'
    ]

    for pname in possible_names:
        val = safe_get_double_param_by_name(duct, pname, None)
        if val is not None:
            return val

    return 0.0


def safe_set_string_param_by_name(element, param_name, value):
    try:
        if element is None or not param_name:
            return False

        if value is None:
            return False

        value = str(value).strip()
        if not value:
            return False

        param = element.LookupParameter(param_name)
        if not param or param.IsReadOnly:
            return False

        if param.StorageType == DB.StorageType.String:
            param.Set(value)
            return True

        set_parameter_by_name(element, param_name, value)
        return True
    except:
        return False


def safe_copy_param_as_string(source, target, source_param_name, target_param_name):
    try:
        if source is None or target is None:
            return False

        value = get_parameter_value_by_name_AsString(source, source_param_name)
        if value is None:
            return False

        value = str(value).strip()
        if not value:
            return False

        return safe_set_string_param_by_name(target, target_param_name, value)
    except:
        return False


def copy_fab_params(source, target):
    safe_copy_param_as_string(source, target, 'Overall Size', 'FP_Product Entry')
    safe_copy_param_as_string(source, target, 'Fabrication Service Name', 'FP_Service Name')
    safe_copy_param_as_string(source, target, 'Fabrication Service Abbreviation', 'FP_Service Abbreviation')


def get_level_plane_z_internal(level):
    """
    Converts level project elevation into internal coordinates.
    """
    try:
        pt = PointConverter(0.0, 0.0, level.ProjectElevation, 'project', doc).internal
        return pt.Z
    except:
        try:
            return level.Elevation
        except:
            return 0.0


def get_upper_level(view, all_levels):
    gen_level = view.GenLevel
    if not gen_level:
        return None

    base_z = get_level_plane_z_internal(gen_level)
    candidates = []

    for lvl in all_levels:
        try:
            if get_level_plane_z_internal(lvl) > base_z:
                candidates.append(lvl)
        except:
            pass

    if not candidates:
        return None

    return min(candidates, key=lambda l: get_level_plane_z_internal(l))


def point_in_active_section_box(view, point):
    if view.ViewType != ViewType.ThreeD:
        return True

    try:
        if not view.IsSectionBoxActive:
            return True

        sb = view.GetSectionBox()
        local_pt = sb.Transform.Inverse.OfPoint(point)

        return (
            sb.Min.X <= local_pt.X <= sb.Max.X and
            sb.Min.Y <= local_pt.Y <= sb.Max.Y and
            sb.Min.Z <= local_pt.Z <= sb.Max.Z
        )
    except:
        return True


def align_instance_top_z_to_point(inst, target_pt):
    """
    Moves the instance vertically so the top of its geometry aligns to target_pt.Z.
    This matches families modeled downward from Ref. Level.
    """
    doc.Regenerate()

    bbox = inst.get_BoundingBox(None)
    if not bbox:
        return

    dz = target_pt.Z - bbox.Max.Z
    if abs(dz) > 1e-6:
        ElementTransformUtils.MoveElement(doc, inst.Id, XYZ(0, 0, dz))
        doc.Regenerate()


# --------------------------------------------------
# File-based settings
# --------------------------------------------------
temp_folder = r"c:\Temp"

if not os.path.exists(temp_folder):
    os.makedirs(temp_folder)

sleeve_length_file = os.path.join(temp_folder, 'Ribbon_Sleeve.txt')
if not os.path.exists(sleeve_length_file):
    with open(sleeve_length_file, 'w') as f:
        f.write('6')

with open(sleeve_length_file, 'r') as f:
    SleeveLength = float(f.read().strip())

annular_file = os.path.join(temp_folder, 'Ribbon_Duct-Floor-Sleeve.txt')
if not os.path.exists(annular_file):
    with open(annular_file, 'w') as f:
        f.write('1')

with open(annular_file, 'r') as f:
    AnnularSpace = float(f.read().strip()) / 12.0


# --------------------------------------------------
# Family loading
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


class SleeveFamilyManager(object):
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
            try:
                if fam.Name.strip().upper() == self.family_name.strip().upper():
                    self.family = fam
                    return fam
            except:
                pass

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
            raise Exception("Family file not found:\n{}".format(self.family_path))

        t = None
        loaded_family_ref = clr.Reference[Family]()

        uiapp.DialogBoxShowing += shared_family_dialog_fallback
        try:
            t = Transaction(self.doc, "Load {} Family".format(self.family_name))
            t.Start()

            result = self.doc.LoadFamily(
                self.family_path,
                FamilyLoaderOptionsHandler(),
                loaded_family_ref
            )

            t.Commit()

            if loaded_family_ref.Value and loaded_family_ref.Value.IsValidObject:
                self.family = loaded_family_ref.Value
                return self.family

            fam = self.get_family_by_name()
            if fam:
                self.family = fam
                return fam

            raise Exception("LoadFamily returned {}, but '{}' was not found.".format(result, self.family_name))

        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise Exception("Family load error:\n{}".format(str(e)))

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

    def get_symbol_by_type_name(self, type_name=None):
        if not self.symbol_cache:
            self.build_symbol_cache()

        if type_name:
            sym = self.symbol_cache.get(type_name.strip().upper())
            if sym:
                return sym

        for k in sorted(self.symbol_cache.keys()):
            return self.symbol_cache[k]

        return None

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

    def get_ready_symbol(self, preferred_type_name=None):
        fam = self.load_family_if_missing()
        if not fam:
            return None

        self.build_symbol_cache()
        sym = self.get_symbol_by_type_name(preferred_type_name)

        if not sym:
            raise Exception(
                "No usable type found in family '{}'. Available types: {}".format(
                    self.family_name,
                    ", ".join(sorted(self.symbol_cache.keys())) if self.symbol_cache else "<none>"
                )
            )

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


# --------------------------------------------------
# Duct helpers
# --------------------------------------------------
def is_fabrication_duct(el):
    try:
        if el is None:
            return False
        if el.Category is None:
            return False
        return el.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationDuctwork)
    except:
        return False


def get_connectors(el):
    try:
        return list(el.ConnectorManager.Connectors)
    except:
        return []


def get_shape_connector(duct):
    connectors = get_connectors(duct)

    for conn in connectors:
        try:
            if conn.Shape in [DB.ConnectorProfileType.Round, DB.ConnectorProfileType.Rectangular]:
                return conn
        except:
            pass

    return None


def is_vertical_duct(duct):
    try:
        loc = duct.Location
        if isinstance(loc, LocationCurve):
            curve = loc.Curve
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            vec = (p1 - p0).Normalize()
            return abs(vec.Z) > 0.99
    except:
        pass

    connectors = get_connectors(duct)
    if len(connectors) < 2:
        return False

    try:
        p0 = connectors[0].Origin
        p1 = connectors[1].Origin
        vec = (p1 - p0).Normalize()
        return abs(vec.Z) > 0.99
    except:
        return False


def get_duct_intersections(duct, level):
    if not is_vertical_duct(duct):
        return []

    bbox = duct.get_BoundingBox(None)
    if bbox is None:
        return []

    plane_z_internal = get_level_plane_z_internal(level)

    if not (bbox.Min.Z < plane_z_internal < bbox.Max.Z):
        return []

    try:
        loc = duct.Location
        if isinstance(loc, LocationCurve):
            curve = loc.Curve
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)

            if abs(p1.Z - p0.Z) > 1e-9:
                t = (plane_z_internal - p0.Z) / (p1.Z - p0.Z)
                x = p0.X + (p1.X - p0.X) * t
                y = p0.Y + (p1.Y - p0.Y) * t
                return [XYZ(x, y, plane_z_internal)]
    except:
        pass

    connectors = get_connectors(duct)
    if len(connectors) >= 2:
        try:
            cx = (connectors[0].Origin.X + connectors[1].Origin.X) / 2.0
            cy = (connectors[0].Origin.Y + connectors[1].Origin.Y) / 2.0
            return [XYZ(cx, cy, plane_z_internal)]
        except:
            pass

    return []


def get_shape_and_size_from_connector(connector, duct):
    if not connector:
        raise Exception("Could not determine duct connector shape.")

    insulation_thickness = get_insulation_thickness(duct)

    if connector.Shape == DB.ConnectorProfileType.Round:
        duct_diameter = connector.Radius * 2.0
        sleeve_diameter = duct_diameter + (2.0 * insulation_thickness) + AnnularSpace

        return "ROUND", {
            "Diameter": sleeve_diameter
        }

    elif connector.Shape == DB.ConnectorProfileType.Rectangular:
        duct_width = connector.Width
        duct_height = connector.Height

        sleeve_width = duct_width + (2.0 * insulation_thickness) + AnnularSpace
        sleeve_height = duct_height + (2.0 * insulation_thickness) + AnnularSpace

        return "RECTANGULAR", {
            "Width": sleeve_width,
            "Height": sleeve_height
        }

    raise Exception("Only round and rectangular fabrication ductwork are supported.")


def get_rectangular_rotation_angle(duct):
    connectors = get_connectors(duct)

    for conn in connectors:
        try:
            if conn.Shape != DB.ConnectorProfileType.Rectangular:
                continue

            basis = conn.CoordinateSystem.BasisX
            vec = XYZ(basis.X, basis.Y, 0.0)

            if vec.GetLength() < 1e-9:
                basis = conn.CoordinateSystem.BasisY
                vec = XYZ(basis.X, basis.Y, 0.0)

            if vec.GetLength() < 1e-9:
                continue

            angle = math.atan2(vec.Y, vec.X)
            angle += math.radians(RECTANGULAR_ROTATION_OFFSET_DEGREES)
            return angle
        except:
            pass

    return math.radians(RECTANGULAR_ROTATION_OFFSET_DEGREES)


def get_family_name(instance):
    try:
        sym = doc.GetElement(instance.GetTypeId())
        if sym and sym.Family:
            return sym.Family.Name
    except:
        pass
    return ""


def is_duplicate_sleeve(point, existing, tol=0.02):
    for el in existing:
        try:
            loc = getattr(el.Location, 'Point', None)
            if loc and \
               abs(loc.X - point.X) < tol and \
               abs(loc.Y - point.Y) < tol and \
               abs(loc.Z - point.Z) < tol:
                return True
        except:
            pass
    return False


def zero_instance_elevation_params(inst):
    for pname in ['Offset', 'Elevation from Level', 'Middle Elevation']:
        try:
            p = inst.LookupParameter(pname)
            if p and not p.IsReadOnly and p.StorageType == DB.StorageType.Double:
                p.Set(0.0)
        except:
            pass


def place_sleeve_at_intersection(duct, pt, level, round_symbol, rect_symbol, existing):
    if is_duplicate_sleeve(pt, existing):
        return None

    connector = get_shape_connector(duct)
    if not connector:
        return None

    shape_name, size_data = get_shape_and_size_from_connector(connector, duct)

    if shape_name == "ROUND":
        symbol = round_symbol
    elif shape_name == "RECTANGULAR":
        symbol = rect_symbol
    else:
        return None

    inst = doc.Create.NewFamilyInstance(
        pt,
        symbol,
        level,
        DB.Structure.StructuralType.NonStructural
    )

    if shape_name == "ROUND":
        set_parameter_by_name(inst, 'Diameter', size_data['Diameter'])

    elif shape_name == "RECTANGULAR":
        set_parameter_by_name(inst, 'Width', size_data['Width'])
        set_parameter_by_name(inst, 'Height', size_data['Height'])

    set_parameter_by_name(inst, 'Length', SleeveLength)

    schedule_level_param = inst.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly:
        try:
            schedule_level_param.Set(level.Id)
        except:
            pass

    zero_instance_elevation_params(inst)
    copy_fab_params(duct, inst)

    if shape_name == "RECTANGULAR":
        align_instance_top_z_to_point(inst, pt)

        angle = get_rectangular_rotation_angle(duct)
        if abs(angle) > 1e-9:
            axis = Line.CreateBound(pt, XYZ(pt.X, pt.Y, pt.Z + 1.0))
            ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle)

    return inst


# --------------------------------------------------
# Main
# --------------------------------------------------
def main():
    round_manager = SleeveFamilyManager(doc, ROUND_FAMILY_NAME, round_family_path)
    rect_manager = SleeveFamilyManager(doc, RECT_FAMILY_NAME, rect_family_path)

    round_symbol = round_manager.get_ready_symbol(ROUND_TYPE_NAME)
    rect_symbol = rect_manager.get_ready_symbol(RECT_TYPE_NAME)

    if not round_symbol:
        show_message("Error", "Cannot load round floor sleeve family.")
        return

    if not rect_symbol:
        show_message("Error", "Cannot load rectangular floor sleeve family.")
        return

    selected_ids = uidoc.Selection.GetElementIds()
    all_levels = list(FilteredElementCollector(doc).OfClass(DB.Level).ToElements())

    existing_sleeves = []
    for el in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType():
        fam_name = get_family_name(el)
        if fam_name in [ROUND_FAMILY_NAME, RECT_FAMILY_NAME]:
            existing_sleeves.append(el)

    is_3d = curview.ViewType == ViewType.ThreeD
    is_plan = curview.ViewType == ViewType.FloorPlan

    if not is_3d and not is_plan:
        show_message("Error", "This script supports 3D and Floor Plan views only.")
        return

    visible_ducts = list(
        FilteredElementCollector(doc, curview.Id)
        .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    if selected_ids.Count > 0:
        ducts = []
        for eid in selected_ids:
            el = doc.GetElement(eid)
            if is_fabrication_duct(el):
                ducts.append(el)
    else:
        ducts = visible_ducts

    if not ducts:
        show_message("Result", "No fabrication ductwork found.")
        return

    if is_plan:
        upper = get_upper_level(curview, all_levels)
        if not upper:
            show_message("Error", "No level above current floor plan view.")
            return
        levels_to_check = [upper]
    else:
        levels_to_check = all_levels

    t = None
    try:
        t = Transaction(doc, 'Place Duct Floor Sleeves at Level Intersections')
        t.Start()

        placed = 0

        for duct in ducts:
            if not is_fabrication_duct(duct):
                continue

            if not is_vertical_duct(duct):
                continue

            for lvl in levels_to_check:
                pts = get_duct_intersections(duct, lvl)

                for pt in pts:
                    if not point_in_active_section_box(curview, pt):
                        continue

                    sleeve = place_sleeve_at_intersection(
                        duct,
                        pt,
                        lvl,
                        round_symbol,
                        rect_symbol,
                        existing_sleeves
                    )

                    if sleeve:
                        placed += 1
                        existing_sleeves.append(sleeve)

        t.Commit()
        show_message("Result", "Placed {} sleeve instances.".format(placed))

    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        show_message("Error", str(e))

    finally:
        if t:
            t.Dispose()


if __name__ == '__main__':
    main()