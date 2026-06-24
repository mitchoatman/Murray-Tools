# -*- coding: utf-8 -*-
import clr
import sys
import os
import re
from math import atan2
from fractions import Fraction

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    LocationCurve,
    Transaction,
    Wall
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException

from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

FamilyName = 'RDS'
FamilyType = 'RDS'

try:
    active_view_level = active_view.GenLevel
except:
    active_view_level = None


# --------------------------------------------------
# Family load options / manager
# --------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


class RDSFamilyManager(object):
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

    def load_family_if_missing(self):
        fam = self.get_family_by_name()
        if fam:
            return fam

        if not os.path.exists(self.family_path):
            TaskDialog.Show("Error", "Family file not found:\n{}".format(self.family_path))
            return None

        t = None
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
            TaskDialog.Show("Error", "Family load error:\n{}".format(str(e)))
            return None
        finally:
            if t:
                t.Dispose()

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
            TaskDialog.Show("Error", "Family symbol activation error:\n{}".format(str(e)))
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


# --------------------------------------------------
# Settings / family path
# --------------------------------------------------
path, filename = os.path.split(__file__)
family_filename = 'RDS.rfa'
family_pathCC = os.path.join(path, family_filename)

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Duct-Wall-Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

try:
    with open(filepath, 'r') as f:
        AnnularSpace = float(f.read())
except:
    AnnularSpace = 1.0
    TaskDialog.Show("Warning", "Failed to read annular space. Using default 1 inch.")

AnnularSpace = AnnularSpace / 12.0

family_manager = RDSFamilyManager(doc, FamilyName, family_pathCC)
famsymb = family_manager.get_ready_symbol(FamilyType)

if not famsymb:
    sys.exit(1)


# --------------------------------------------------
# Selection filters
# --------------------------------------------------
class FabricationDuctSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            if elem is None:
                return False
            if not isinstance(elem, DB.FabricationPart):
                return False
            if elem.Category is None:
                return False

            fab_duct_cat_id = DB.ElementId(DB.BuiltInCategory.OST_FabricationDuctwork).IntegerValue
            return elem.Category.Id.IntegerValue == fab_duct_cat_id
        except:
            return False

    def AllowReference(self, reference, point):
        return False


class LinkedWallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.RevitLinkInstance)

    def AllowReference(self, reference, position):
        try:
            link_instance = doc.GetElement(reference.ElementId)
            if not isinstance(link_instance, DB.RevitLinkInstance):
                return False

            linked_doc = link_instance.GetLinkDocument()
            if linked_doc is None:
                return False

            linked_elem = linked_doc.GetElement(reference.LinkedElementId)
            return isinstance(linked_elem, DB.Wall)
        except:
            return False


# --------------------------------------------------
# Helpers
# --------------------------------------------------
def get_best_level_id(element):
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


def safe_set_double_param_if_exists(element, param_name, value):
    try:
        p = element.LookupParameter(param_name)
        if p and not p.IsReadOnly and p.StorageType == DB.StorageType.Double:
            p.Set(value)
            return True
    except:
        pass
    return False


def select_fabrication_duct():
    try:
        duct_ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            FabricationDuctSelectionFilter(),
            "Select a round MEP Fabrication Duct"
        )
        return doc.GetElement(duct_ref.ElementId)
    except:
        return None


def select_linked_wall():
    return uidoc.Selection.PickObject(
        ObjectType.LinkedElement,
        LinkedWallSelectionFilter(),
        "Select a wall in a Revit link"
    )


def get_pipe_centerline(duct):
    pipe_location = duct.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    raise Exception("Invalid duct centerline")


def get_connectors(element):
    try:
        return list(element.ConnectorManager.Connectors)
    except:
        raise Exception("Could not read connectors from selected duct.")


def frac2string(s):
    i, f = s.groups(0)
    f = Fraction(f)
    return str(int(i) + float(f))


def get_diameter_from_duct(duct, connector=None):
    try:
        if connector and connector.Shape == DB.ConnectorProfileType.Round:
            return connector.Radius * 2.0
    except:
        pass

    overall_size = get_parameter_value_by_name_AsString(duct, 'Overall Size')
    if not overall_size:
        raise Exception("Could not read duct Overall Size.")

    if '/' in overall_size:
        return float(
            re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)[^\d.]', frac2string, overall_size)
        ) / 12.0
    else:
        return float(re.sub(r'[^\d.]', '', overall_size)) / 12.0


def get_linked_wall_thickness_and_curve(link_ref):
    link_instance = doc.GetElement(link_ref.ElementId)
    if not link_instance:
        raise Exception("Could not get Revit link instance.")

    linked_doc = link_instance.GetLinkDocument()
    if linked_doc is None:
        raise Exception("Could not access linked document.")

    wall = linked_doc.GetElement(link_ref.LinkedElementId)
    if not isinstance(wall, Wall):
        raise Exception("Selected linked element is not a wall.")

    wall_type = wall.WallType
    thickness = wall_type.get_Parameter(DB.BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()

    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        raise Exception("Selected wall does not have a valid location curve.")

    wall_curve = loc.Curve

    try:
        transform = link_instance.GetTotalTransform()
    except:
        transform = link_instance.GetTransform()

    transformed_wall_curve = wall_curve.CreateTransformed(transform)
    return thickness, transformed_wall_curve


def project_wall_curve_to_duct_plane(wall_curve, duct_curve):
    duct_start = duct_curve.GetEndPoint(0)
    duct_end = duct_curve.GetEndPoint(1)
    duct_direction = (duct_end - duct_start).Normalize()

    plane_normal = DB.XYZ(0, 0, 1)
    if abs(duct_direction.Z) > 0.99:
        plane_normal = DB.XYZ(1, 0, 0)

    plane = DB.Plane.CreateByNormalAndOrigin(plane_normal, duct_start)

    wall_start = wall_curve.GetEndPoint(0)
    wall_end = wall_curve.GetEndPoint(1)

    uv_start, dist_start = plane.Project(wall_start)
    uv_end, dist_end = plane.Project(wall_end)

    origin = plane.Origin
    x_axis = plane.XVec
    y_axis = plane.YVec

    projected_start = origin + x_axis * uv_start.U + y_axis * uv_start.V
    projected_end = origin + x_axis * uv_end.U + y_axis * uv_end.V

    return DB.Line.CreateBound(projected_start, projected_end)


def get_intersection_point(duct_curve, wall_curve):
    result_array = clr.Reference[DB.IntersectionResultArray]()
    result = duct_curve.Intersect(wall_curve, result_array)

    if result != DB.SetComparisonResult.Overlap:
        raise Exception("Duct does not intersect the selected linked wall.")

    if result_array.Value and result_array.Value.Size > 0:
        return result_array.Value[0].XYZPoint

    raise Exception("Failed to retrieve intersection point.")


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


def rotate_to_duct(instance, connectors, insertion_point, reference_point):
    if len(connectors) < 2:
        return

    conn1, conn2 = connectors[0], connectors[1]
    nearest_conn = min([conn1, conn2], key=lambda c: reference_point.DistanceTo(c.Origin))
    other_conn = conn2 if nearest_conn == conn1 else conn1

    vec = other_conn.Origin - nearest_conn.Origin
    angle = atan2(vec.Y, vec.X)

    axis = DB.Line.CreateBound(
        insertion_point,
        DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1.0)
    )

    DB.ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)


# --------------------------------------------------
# Placement
# --------------------------------------------------
def place_and_modify_family(duct, wall_ref, famsymb):
    try:
        centerline_curve = get_pipe_centerline(duct)
        wall_thickness, wall_curve = get_linked_wall_thickness_and_curve(wall_ref)

        projected_wall_curve = project_wall_curve_to_duct_plane(wall_curve, centerline_curve)
        intersection_point = get_intersection_point(centerline_curve, projected_wall_curve)

        pipe_connectors = get_connectors(duct)
        if len(pipe_connectors) < 2:
            raise Exception("Selected duct does not have at least 2 connectors.")

        nearest_connector = min(pipe_connectors, key=lambda c: intersection_point.DistanceTo(c.Origin))
        other_connector = [c for c in pipe_connectors if c.Id != nearest_connector.Id][0]
        duct_direction = (other_connector.Origin - nearest_connector.Origin).Normalize()

        insertion_point = intersection_point - duct_direction * (wall_thickness / 2.0)

        level_id = get_best_level_id(duct)
        level = doc.GetElement(level_id) if level_id else None

        new_family_instance = create_family_instance(insertion_point, famsymb, level_id)
        if not new_family_instance:
            raise Exception("Failed to create family instance.")

        diameter = get_diameter_from_duct(duct, nearest_connector) + AnnularSpace
        set_parameter_by_name(new_family_instance, 'Diameter', diameter)

        safe_set_double_param_if_exists(new_family_instance, 'Length', wall_thickness)

        rotate_to_duct(new_family_instance, pipe_connectors, insertion_point, intersection_point)

        set_parameter_by_name(
            new_family_instance,
            'FP_Service Name',
            get_parameter_value_by_name_AsString(duct, 'Fabrication Service Name')
        )

        schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
        if schedule_level_param and not schedule_level_param.IsReadOnly and level:
            schedule_level_param.Set(level.Id)

        return True

    except Exception as e:
        raise Exception("Family placement error: {}".format(e))


# --------------------------------------------------
# Main
# --------------------------------------------------
t = None
wall_ref = None

try:
    wall_ref = select_linked_wall()

    while True:
        try:
            duct = select_fabrication_duct()
            if not duct:
                break

            if duct.ItemCustomId == 1:
                TaskDialog.Show("Warning", "You selected a rectangular duct.")
                continue

            t = Transaction(doc, 'Place RDS in Linked Wall')
            t.Start()

            if not place_and_modify_family(duct, wall_ref, famsymb):
                t.RollBack()
                break

            t.Commit()

        except OperationCanceledException:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            break

        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            TaskDialog.Show("Error", "Error during operation:\n{}".format(e))
            break

        finally:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            if t:
                t.Dispose()
                t = None

except OperationCanceledException:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()

except Exception as e:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    TaskDialog.Show("Error", "Error during execution:\n{}".format(e))

finally:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    if t:
        t.Dispose()