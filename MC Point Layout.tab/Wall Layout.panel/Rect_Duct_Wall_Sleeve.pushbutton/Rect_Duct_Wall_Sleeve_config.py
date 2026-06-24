# -*- coding: utf-8 -*-
import clr
import os
import sys
from math import atan2

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    FamilySymbol,
    BuiltInCategory,
    LocationCurve,
    Transaction,
    Wall
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsDouble
)
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
app = __revit__.Application
active_view = doc.ActiveView

try:
    active_view_level = active_view.GenLevel
except:
    active_view_level = None


# --------------------------------------------------
# Helpers
# --------------------------------------------------
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


def load_family(family_path):
    t = None
    try:
        t = Transaction(doc, 'Load RWS Family')
        t.Start()

        families = FilteredElementCollector(doc).OfClass(Family)
        family_name = 'RWS'

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


def get_rws_symbol():
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_DuctAccessory
    ).OfClass(FamilySymbol)

    family_name = 'RWS'
    family_type = 'RWS'

    for fs in collector:
        try:
            if fs.Family.Name == family_name:
                type_name = fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                if type_name == family_type:
                    return fs
        except:
            pass

    return None


def activate_symbol(symbol):
    if not symbol:
        raise Exception("Family symbol not found.")

    if symbol.IsActive:
        return

    t = None
    try:
        t = Transaction(doc, 'Activate RWS Symbol')
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
# Selection
# --------------------------------------------------
def select_fabrication_duct():
    try:
        duct_ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            FabricationDuctSelectionFilter(),
            "Select a rectangular fabrication duct"
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


# --------------------------------------------------
# Geometry / sizing
# --------------------------------------------------
def get_duct_centerline(duct):
    loc = duct.Location
    if isinstance(loc, LocationCurve):
        return loc.Curve
    raise Exception("Selected duct does not have a valid centerline.")


def get_connectors(element):
    try:
        return list(element.ConnectorManager.Connectors)
    except:
        raise Exception("Could not read connectors from selected duct.")


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

    wall_loc = wall.Location
    if not isinstance(wall_loc, LocationCurve):
        raise Exception("Selected wall does not have a valid location curve.")

    wall_curve = wall_loc.Curve

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


def get_rectangular_size_from_connector(connector):
    if connector.Shape != DB.ConnectorProfileType.Rectangular:
        raise Exception("RWS family supports rectangular fabrication duct only.")

    return connector.Width, connector.Height


def is_rectangular_fab_duct(duct):
    try:
        connectors = get_connectors(duct)
        for c in connectors:
            if c.Shape == DB.ConnectorProfileType.Rectangular:
                return True
    except:
        pass
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
path, filename = os.path.split(__file__)
family_path = os.path.join(path, 'RWS.rfa')

folder_name = r"c:\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Duct-Wall-Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

with open(filepath, 'r') as f:
    AnnularSpace = float(f.read()) / 12.0


def place_rws_in_linked_wall(duct, wall_ref, famsymb):
    duct_curve = get_duct_centerline(duct)
    wall_thickness, wall_curve = get_linked_wall_thickness_and_curve(wall_ref)

    projected_wall_curve = project_wall_curve_to_duct_plane(wall_curve, duct_curve)
    intersection_point = get_intersection_point(duct_curve, projected_wall_curve)

    connectors = get_connectors(duct)
    if len(connectors) < 2:
        raise Exception("Selected duct does not have enough connectors.")

    nearest_conn = min(connectors, key=lambda c: intersection_point.DistanceTo(c.Origin))
    other_conn = [c for c in connectors if c.Id != nearest_conn.Id][0]
    duct_direction = (other_conn.Origin - nearest_conn.Origin).Normalize()

    insertion_point = intersection_point - duct_direction * (wall_thickness / 2.0)

    width, height = get_rectangular_size_from_connector(nearest_conn)

    level_id = safe_get_level_id(duct)
    new_family_instance = create_family_instance(insertion_point, famsymb, level_id)

    if not new_family_instance:
        raise Exception("Failed to create family instance.")

    set_parameter_by_name(new_family_instance, 'Width', width + AnnularSpace)
    set_parameter_by_name(new_family_instance, 'Height', height + AnnularSpace)

    safe_set_length_param_if_exists(new_family_instance, 'Length', wall_thickness)

    rotate_to_duct(new_family_instance, connectors, insertion_point, intersection_point)

    try:
        set_parameter_by_name(
            new_family_instance,
            'FP_Service Name',
            get_parameter_value_by_name_AsString(duct, 'Fabrication Service Name')
        )
    except:
        pass

    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly and level_id:
        schedule_level_param.Set(level_id)


# --------------------------------------------------
# Main
# --------------------------------------------------
selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

if selection:
    t = None
    try:
        t = Transaction(doc, "Update FP Parameters")
        t.Start()

        for x in selection:
            try:
                oldheight = get_parameter_value_by_name_AsDouble(x, 'Height')
                oldwidth = get_parameter_value_by_name_AsDouble(x, 'Width')
                set_parameter_by_name(x, 'Height', oldwidth)
                set_parameter_by_name(x, 'Width', oldheight)
            except:
                continue

        t.Commit()

    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        show_message("Error", "Parameter update error:\n{}".format(e))

    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()

else:
    if not os.path.exists(family_path):
        show_message("Error", "Family file not found:\n{}".format(family_path))
        sys.exit()

    try:
        load_family(family_path)
    except Exception as e:
        show_message("Error", str(e))
        sys.exit()

    famsymb = get_rws_symbol()
    if not famsymb:
        show_message("Error", "Could not find family type 'RWS' in family 'RWS'.")
        sys.exit()

    try:
        activate_symbol(famsymb)
    except Exception as e:
        show_message("Error", str(e))
        sys.exit()

    t = None
    wall_ref = None

    try:
        wall_ref = select_linked_wall()

        while True:
            try:
                duct = select_fabrication_duct()
                if not duct:
                    break

                if not is_rectangular_fab_duct(duct):
                    show_message("Warning", "You selected a non-rectangular duct.")
                    continue

                t = Transaction(doc, 'Place RWS in Linked Wall')
                t.Start()

                place_rws_in_linked_wall(duct, wall_ref, famsymb)

                t.Commit()

            except OperationCanceledException:
                if t and t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                break

            except Exception as e:
                if t and t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                show_message("Error", "Error during operation:\n{}".format(str(e)))
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
        show_message("Error", "Error during execution:\n{}".format(str(e)))

    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()