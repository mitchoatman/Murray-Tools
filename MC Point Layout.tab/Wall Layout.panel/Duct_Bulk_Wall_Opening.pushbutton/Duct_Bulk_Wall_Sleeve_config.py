# -*- coding: utf-8 -*-
import clr
import os
import sys
import re
from math import atan2
from fractions import Fraction

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    ElementId,
    Family,
    FamilySymbol,
    FilteredElementCollector,
    Line,
    LocationCurve,
    RevitLinkInstance,
    Transaction,
    Wall,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
)
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

RECT_SERVICE_TYPE_NAME = "Rectangular Duct"
ROUND_SERVICE_TYPE_NAME = "Round Duct"
FIRE_DAMPER_SERVICE_TYPE_NAME = "Fire Damper"

FIRE_SMOKE_DAMPER_KEY = "Fire Smoke Damper"

RECT_FAMILY_NAME = 'RWS'
RECT_FAMILY_TYPE = 'RWS'
ROUND_FAMILY_NAME = 'RDS'
ROUND_FAMILY_TYPE = 'RDS'

try:
    active_view_level = active_view.GenLevel
except:
    active_view_level = None

try:
    fabrication_config = DB.FabricationConfiguration.GetFabricationConfiguration(doc)
except:
    fabrication_config = None


# --------------------------------------------------
# Helpers
# --------------------------------------------------

def get_annular_space_for_element(elem):
    service_type_name = get_fab_part_service_type_name(elem)

    if service_type_name == FIRE_DAMPER_SERVICE_TYPE_NAME:
        service_name = FIRE_SMOKE_DAMPER_KEY
    else:
        service_name = get_service_name_for_annular(elem)

    annular_inches = service_annular_map.get(service_name, DEFAULT_ANNULAR_INCHES)
    annular_feet = annular_inches / 12.0
    return annular_feet, service_name


def get_service_name_for_annular(element):
    param_names = [
        "Fabrication Service Name",
        "Service Name",
        "System Name",
        "System Abbreviation",
        "Abbreviation"
    ]

    for pname in param_names:
        try:
            p = element.LookupParameter(pname)
            if p:
                val = p.AsString() or p.AsValueString()
                if val and val.strip():
                    return val.strip()
        except:
            pass

    try:
        val = get_parameter_value_by_name_AsString(element, "Fabrication Service Name")
        if val and val.strip():
            return val.strip()
    except:
        pass

    return "UNASSIGNED"


def get_fab_part_service_type_name(part):
    # Preferred: direct API property on FabricationPart
    try:
        if isinstance(part, DB.FabricationPart) and fabrication_config:
            service_type_name = fabrication_config.GetServiceTypeName(part.ServiceType)
            if service_type_name and service_type_name.strip():
                return service_type_name.strip()
    except:
        pass

    # Fallbacks only if direct API lookup fails
    param_names = [
        "Item Service Type",
        "Service Type",
        "FP_Service Type",
    ]

    for pname in param_names:
        try:
            val = get_parameter_value_by_name_AsString(part, pname)
            if val and val.strip():
                return val.strip()
        except:
            pass

        try:
            p = part.LookupParameter(pname)
            if p:
                val = p.AsString() or p.AsValueString()
                if val and val.strip():
                    return val.strip()
        except:
            pass

    return None


def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        print("{}: {}".format(title, message))


def safe_get_level_id(element):
    try:
        if element.LevelId and element.LevelId != DB.ElementId.InvalidElementId:
            return element.LevelId
    except:
        pass

    try:
        if active_view_level:
            return active_view_level.Id
    except:
        pass

    return None


def safe_set_length_param_if_exists(element, param_name, value):
    try:
        p = element.LookupParameter(param_name)
        if p and not p.IsReadOnly and p.StorageType == DB.StorageType.Double:
            p.Set(value)
            return True
    except:
        pass
    return False


def get_fabrication_service_name(element):
    try:
        p = element.LookupParameter("Fabrication Service Name")
        if p:
            val = p.AsValueString()
            if val:
                return val
    except:
        pass

    try:
        val = get_parameter_value_by_name_AsString(element, "Fabrication Service Name")
        if val:
            return val
    except:
        pass

    try:
        val = get_parameter_value_by_name_AsString(element, "Fabrication Service")
        if val:
            return val
    except:
        pass

    return None


def get_connectors(element):
    try:
        return list(element.ConnectorManager.Connectors)
    except:
        raise Exception("Could not read connectors from selected element.")


def get_element_shape(elem):
    service_type_name = get_fab_part_service_type_name(elem)

    if service_type_name == RECT_SERVICE_TYPE_NAME:
        return "RECT"
    elif service_type_name == ROUND_SERVICE_TYPE_NAME:
        return "ROUND"
    elif service_type_name == FIRE_DAMPER_SERVICE_TYPE_NAME:
        connectors = get_connectors(elem)
        for c in connectors:
            try:
                if c.Shape == DB.ConnectorProfileType.Rectangular:
                    return "RECT"
                elif c.Shape == DB.ConnectorProfileType.Round:
                    return "ROUND"
            except:
                pass

    return None


def is_supported_fab_element(elem):
    try:
        if elem is None:
            return False
        if not isinstance(elem, DB.FabricationPart):
            return False
        return get_element_shape(elem) in ["RECT", "ROUND"]
    except:
        return False


def get_shape_connectors(elem, shape_name):
    connectors = get_connectors(elem)

    if shape_name == "RECT":
        target_shape = DB.ConnectorProfileType.Rectangular
    elif shape_name == "ROUND":
        target_shape = DB.ConnectorProfileType.Round
    else:
        return []

    return [c for c in connectors if c.Shape == target_shape]


def get_farthest_connector_pair(connectors):
    if len(connectors) < 2:
        return None, None

    best_pair = (None, None)
    best_dist = -1.0

    for i in range(len(connectors)):
        for j in range(i + 1, len(connectors)):
            try:
                dist = connectors[i].Origin.DistanceTo(connectors[j].Origin)
                if dist > best_dist:
                    best_dist = dist
                    best_pair = (connectors[i], connectors[j])
            except:
                pass

    return best_pair


def get_element_centerline(elem, shape_name):
    loc = elem.Location
    if isinstance(loc, LocationCurve):
        return loc.Curve

    shape_connectors = get_shape_connectors(elem, shape_name)
    c1, c2 = get_farthest_connector_pair(shape_connectors)

    if not c1 or not c2:
        raise Exception("Element does not have two usable connectors for centerline.")

    return Line.CreateBound(c1.Origin, c2.Origin)


def get_rectangular_size_from_connector(connector):
    if connector.Shape != DB.ConnectorProfileType.Rectangular:
        raise Exception("RWS family supports rectangular fabrication duct only.")
    return connector.Width, connector.Height


def frac2string(match):
    whole = match.group(1)
    frac = match.group(2)

    val = float(Fraction(frac))
    if whole:
        val += int(whole)

    return str(val)


def get_round_diameter_from_duct(elem, connector=None):
    try:
        if connector and connector.Shape == DB.ConnectorProfileType.Round:
            return connector.Radius * 2.0
    except:
        pass

    overall_size = get_parameter_value_by_name_AsString(elem, 'Overall Size')
    if not overall_size:
        raise Exception("Could not read element Overall Size.")

    size_text = overall_size.strip()

    if '/' in size_text:
        size_text = re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)', frac2string, size_text)

    numeric = re.sub(r'[^\d.]', '', size_text)
    if not numeric:
        raise Exception("Could not parse round diameter from Overall Size.")

    return float(numeric) / 12.0


def get_linked_wall_thickness_and_curve(link_instance, wall):
    if not isinstance(wall, Wall):
        raise Exception("Linked element is not a wall.")

    wall_type = wall.WallType
    thickness_param = wall_type.get_Parameter(DB.BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
    if not thickness_param:
        raise Exception("Could not get wall thickness.")

    thickness = thickness_param.AsDouble()

    wall_loc = wall.Location
    if not isinstance(wall_loc, LocationCurve):
        raise Exception("Wall does not have a valid location curve.")

    wall_curve = wall_loc.Curve

    try:
        transform = link_instance.GetTotalTransform()
    except:
        transform = link_instance.GetTransform()

    transformed_wall_curve = wall_curve.CreateTransformed(transform)
    return thickness, transformed_wall_curve


def get_all_walls_in_link(link_instance):
    linked_doc = link_instance.GetLinkDocument()
    if linked_doc is None:
        return []

    return list(
        FilteredElementCollector(linked_doc)
        .OfClass(Wall)
        .WhereElementIsNotElementType()
    )


# --------------------------------------------------
# Intersection logic
# --------------------------------------------------
def flatten_point_xy(pt):
    return DB.XYZ(pt.X, pt.Y, 0.0)


def flatten_curve_to_xy(curve):
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    return Line.CreateBound(flatten_point_xy(p0), flatten_point_xy(p1))


def get_plan_intersection_point_on_duct(duct_curve, wall_curve):
    flat_duct = flatten_curve_to_xy(duct_curve)
    flat_wall = flatten_curve_to_xy(wall_curve)

    result_array = clr.Reference[DB.IntersectionResultArray]()
    result = flat_duct.Intersect(flat_wall, result_array)

    if result != DB.SetComparisonResult.Overlap:
        return None

    if not result_array.Value or result_array.Value.Size == 0:
        return None

    plan_pt = result_array.Value[0].XYZPoint

    d0 = duct_curve.GetEndPoint(0)
    d1 = duct_curve.GetEndPoint(1)
    flat_d0 = flatten_point_xy(d0)
    flat_d1 = flatten_point_xy(d1)

    flat_dir = flat_d1 - flat_d0
    flat_len = flat_dir.GetLength()
    if flat_len < 1e-9:
        return None

    flat_dir = flat_dir.Normalize()
    proj_dist = (plan_pt - flat_d0).DotProduct(flat_dir)

    if proj_dist < 0.0 or proj_dist > flat_len:
        return None

    real_dir = (d1 - d0).Normalize()
    return d0 + real_dir * proj_dist


def points_close_xy(pt1, pt2, tol=0.25):
    dx = pt1.X - pt2.X
    dy = pt1.Y - pt2.Y
    dist_xy = (dx * dx + dy * dy) ** 0.5
    return dist_xy <= tol


def is_duplicate_intersection(elem_id, intersection_point, placed_keys, tol=0.25):
    for placed_elem_id, placed_pt in placed_keys:
        if placed_elem_id != elem_id:
            continue
        if points_close_xy(intersection_point, placed_pt, tol):
            return True
    return False


def create_family_instance(insertion_point, symbol, level_id=None):
    try:
        if level_id and level_id != DB.ElementId.InvalidElementId:
            lvl = doc.GetElement(level_id)
            if lvl:
                return doc.Create.NewFamilyInstance(
                    insertion_point,
                    symbol,
                    lvl,
                    DB.Structure.StructuralType.NonStructural
                )
    except:
        pass

    return doc.Create.NewFamilyInstance(
        insertion_point,
        symbol,
        DB.Structure.StructuralType.NonStructural
    )


def rotate_instance_to_duct(instance, duct_curve, insertion_point):
    start = duct_curve.GetEndPoint(0)
    end = duct_curve.GetEndPoint(1)

    raw_vec = (end - start).Normalize()
    vec = get_consistent_xy_direction(raw_vec)

    angle = atan2(vec.Y, vec.X)

    axis = DB.Line.CreateBound(
        insertion_point,
        DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1.0)
    )

    DB.ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)


# --------------------------------------------------
# Dynamic centering fix
# --------------------------------------------------
def xyz_scale(v, s):
    return DB.XYZ(v.X * s, v.Y * s, v.Z * s)


def xyz_negate(v):
    return DB.XYZ(-v.X, -v.Y, -v.Z)


def get_consistent_xy_direction(vec):
    v = DB.XYZ(vec.X, vec.Y, 0.0)

    if v.GetLength() < 1e-9:
        return vec.Normalize()

    v = v.Normalize()

    if v.X < -1e-9 or (abs(v.X) <= 1e-9 and v.Y < 0):
        v = xyz_negate(v)

    return v


def get_bbox_center(element):
    bbox = None
    try:
        bbox = element.get_BoundingBox(None)
    except:
        pass

    if not bbox:
        return None

    return DB.XYZ(
        (bbox.Min.X + bbox.Max.X) / 2.0,
        (bbox.Min.Y + bbox.Max.Y) / 2.0,
        (bbox.Min.Z + bbox.Max.Z) / 2.0
    )


def get_instance_long_axis(instance, reference_direction):
    candidates = []

    try:
        tr = instance.GetTransform()
        candidates.extend([tr.BasisX, tr.BasisY, tr.BasisZ])
    except:
        pass

    try:
        candidates.extend([instance.HandOrientation, instance.FacingOrientation])
    except:
        pass

    ref = reference_direction.Normalize()
    best_vec = None
    best_dot = -1.0

    for v in candidates:
        try:
            vn = v.Normalize()
            d = abs(vn.DotProduct(ref))
            if d > best_dot:
                best_dot = d
                best_vec = vn
        except:
            pass

    if best_vec is None:
        return ref

    if best_vec.DotProduct(ref) < 0:
        best_vec = xyz_negate(best_vec)

    return best_vec


def center_instance_on_point_along_axis(instance, target_point, axis):
    try:
        doc.Regenerate()

        bbox_center = get_bbox_center(instance)
        if not bbox_center:
            return False

        axis = axis.Normalize()
        delta = target_point - bbox_center
        move_dist = delta.DotProduct(axis)

        if abs(move_dist) < 1e-6:
            return False

        DB.ElementTransformUtils.MoveElement(
            doc,
            instance.Id,
            xyz_scale(axis, move_dist)
        )

        doc.Regenerate()
        return True
    except:
        return False


# --------------------------------------------------
# Family loading
# --------------------------------------------------
class FamilyLoadOptions(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues[0] = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues[0] = False
        return True


def load_family(family_path, family_name):
    t = None
    try:
        t = Transaction(doc, 'Load {} Family'.format(family_name))
        t.Start()

        families = FilteredElementCollector(doc).OfClass(Family)

        if not any(f.Name == family_name for f in families):
            load_options = FamilyLoadOptions()
            loaded_family = clr.StrongBox[DB.Family]()
            success = doc.LoadFamily(family_path, load_options, loaded_family)
            if not success or not loaded_family.Value:
                raise Exception("Failed to load family '{}'".format(family_name))

        t.Commit()

    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise Exception("Family load error: {}".format(e))

    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()


def get_family_symbol(family_name, family_type):
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)

    for fs in collector:
        try:
            if fs.Family.Name == family_name:
                type_name = fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                if type_name == family_type:
                    return fs
        except:
            pass

    return None


def activate_symbol(symbol, family_name):
    if not symbol:
        raise Exception("Family symbol not found for '{}'.".format(family_name))

    if symbol.IsActive:
        return

    t = None
    try:
        t = Transaction(doc, 'Activate {} Symbol'.format(family_name))
        t.Start()
        symbol.Activate()
        doc.Regenerate()
        t.Commit()
    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise Exception("Family symbol activation error: {}".format(e))
    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()


def get_ready_symbol(family_path, family_name, family_type):
    if not os.path.exists(family_path):
        raise Exception("Family file not found:\n{}".format(family_path))

    load_family(family_path, family_name)

    sym = get_family_symbol(family_name, family_type)
    if not sym:
        raise Exception(
            "Could not find family type '{}' in family '{}'.".format(
                family_type, family_name
            )
        )

    activate_symbol(sym, family_name)
    return sym


# --------------------------------------------------
# Selection filters
# --------------------------------------------------
class FabricationDuctSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, DB.FabricationPart)
        except:
            return False

    def AllowReference(self, reference, point):
        return False


class LinkInstanceSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, RevitLinkInstance)

    def AllowReference(self, reference, point):
        return False


# --------------------------------------------------
# Selection
# --------------------------------------------------
def pick_many_fabrication_elements():
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        FabricationDuctSelectionFilter(),
        "Select fabrication ducts and fire dampers, then click Finish"
    )
    return [doc.GetElement(r.ElementId) for r in refs]


def pick_single_link_instance():
    ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        LinkInstanceSelectionFilter(),
        "Select Revit link instance containing walls"
    )
    return doc.GetElement(ref.ElementId)


# --------------------------------------------------
# Placement settings
# --------------------------------------------------
path, filename = os.path.split(__file__)
rect_family_path = os.path.join(path, 'RWS.rfa')
round_family_path = os.path.join(path, 'RDS.rfa')

folder_name = r"c:\Temp"
filepath = os.path.join(folder_name, "Ribbon_Duct-Wall-Sleeve-Services.txt")
DEFAULT_ANNULAR_INCHES = 0.5


def load_service_settings(path):
    settings = {}
    if not os.path.exists(path):
        return settings

    with open(path, 'r') as f:
        for line in f:
            line = line.strip()
            if "=" not in line:
                continue

            parts = line.split("=", 1)
            if len(parts) != 2:
                continue

            key = parts[0].strip()
            val = parts[1].strip()

            try:
                settings[key] = float(val)
            except:
                pass

    return settings


service_annular_map = load_service_settings(filepath)


def place_sleeve_at_intersection(elem, link_instance, wall, symbol_map, placed_keys):
    shape_name = get_element_shape(elem)
    if shape_name not in symbol_map:
        raise Exception("No family symbol loaded for element shape '{}'.".format(shape_name))

    famsymb = symbol_map[shape_name]

    elem_curve = get_element_centerline(elem, shape_name)
    wall_thickness, wall_curve = get_linked_wall_thickness_and_curve(link_instance, wall)

    intersection_point = get_plan_intersection_point_on_duct(elem_curve, wall_curve)
    if not intersection_point:
        return False

    if is_duplicate_intersection(elem.Id.IntegerValue, intersection_point, placed_keys, 0.25):
        return False

    shape_connectors = get_shape_connectors(elem, shape_name)
    if len(shape_connectors) < 2:
        raise Exception("Selected element does not have enough {} connectors.".format(shape_name.lower()))

    nearest_conn = min(shape_connectors, key=lambda c: intersection_point.DistanceTo(c.Origin))
    other_conn = max(
        [c for c in shape_connectors if c.Id != nearest_conn.Id],
        key=lambda c: c.Origin.DistanceTo(nearest_conn.Origin)
    )

    elem_direction = get_consistent_xy_direction(
        (other_conn.Origin - nearest_conn.Origin).Normalize()
    )

    insertion_point = intersection_point
    level_id = safe_get_level_id(elem)
    annular_feet, service_name = get_annular_space_for_element(elem)
    total_clearance = annular_feet * 2.0

    new_family_instance = create_family_instance(insertion_point, famsymb, level_id)
    if not new_family_instance:
        raise Exception("Failed to create family instance.")

    if shape_name == "RECT":
        width, height = get_rectangular_size_from_connector(nearest_conn)
        set_parameter_by_name(new_family_instance, 'Width', width + total_clearance)
        set_parameter_by_name(new_family_instance, 'Height', height + total_clearance)

    elif shape_name == "ROUND":
        diameter = get_round_diameter_from_duct(elem, nearest_conn)
        set_parameter_by_name(new_family_instance, 'Diameter', diameter + total_clearance)

    safe_set_length_param_if_exists(new_family_instance, 'Length', wall_thickness)

    rotate_instance_to_duct(new_family_instance, elem_curve, insertion_point)

    sleeve_axis = get_instance_long_axis(new_family_instance, elem_direction)
    center_instance_on_point_along_axis(new_family_instance, intersection_point, sleeve_axis)

    try:
        set_parameter_by_name(
            new_family_instance,
            'FP_Service Name',
            get_fabrication_service_name(elem)
        )
    except:
        pass

    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly and level_id:
        schedule_level_param.Set(level_id)

    placed_keys.add((elem.Id.IntegerValue, intersection_point))
    return True


# --------------------------------------------------
# Main
# --------------------------------------------------
try:
    elems = pick_many_fabrication_elements()

    if not elems:
        show_message("Cancelled", "No fabrication elements selected.")
        sys.exit()

    supported_elems = [e for e in elems if is_supported_fab_element(e)]
    skipped_unsupported = len(elems) - len(supported_elems)

    if not supported_elems:
        show_message("Warning", "No supported fabrication ducts or fire dampers were found in the selection.")
        sys.exit()

    shapes_needed = set([get_element_shape(e) for e in supported_elems])

    symbol_map = {}

    if "RECT" in shapes_needed:
        symbol_map["RECT"] = get_ready_symbol(
            rect_family_path,
            RECT_FAMILY_NAME,
            RECT_FAMILY_TYPE
        )

    if "ROUND" in shapes_needed:
        symbol_map["ROUND"] = get_ready_symbol(
            round_family_path,
            ROUND_FAMILY_NAME,
            ROUND_FAMILY_TYPE
        )

    link_instance = pick_single_link_instance()

    if not link_instance:
        show_message("Cancelled", "No Revit link instance selected.")
        sys.exit()

    all_link_walls = get_all_walls_in_link(link_instance)

    if not all_link_walls:
        show_message("Cancelled", "No walls found in selected Revit link.")
        sys.exit()

    placed_count = 0
    placed_rect = 0
    placed_round = 0
    checked_pairs = 0
    skipped_non_intersecting = 0
    error_log = []
    placed_keys = set()

    t = None
    try:
        t = Transaction(doc, 'Batch Place Duct Wall Sleeves')
        t.Start()

        for elem in supported_elems:
            elem_shape = get_element_shape(elem)

            for wall in all_link_walls:
                checked_pairs += 1
                try:
                    placed = place_sleeve_at_intersection(
                        elem,
                        link_instance,
                        wall,
                        symbol_map,
                        placed_keys
                    )

                    if placed:
                        placed_count += 1
                        if elem_shape == "RECT":
                            placed_rect += 1
                        elif elem_shape == "ROUND":
                            placed_round += 1
                    else:
                        skipped_non_intersecting += 1

                except Exception as e:
                    error_log.append(
                        "Element {} / Wall {}: {}".format(
                            elem.Id.IntegerValue,
                            wall.Id.IntegerValue,
                            str(e)
                        )
                    )

        t.Commit()

    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        show_message("Error", "Batch placement error:\n{}".format(str(e)))
        sys.exit()

    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()

except OperationCanceledException:
    show_message("Cancelled", "Operation cancelled by user.")

except Exception as e:
    show_message("Error", "Error during execution:\n{}".format(str(e)))