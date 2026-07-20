# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
import System

from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    LocationCurve,
    Transaction
)
from Autodesk.Revit.Exceptions import OperationCanceledException

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsInteger
)
from Parameters.Add_SharedParameters import Shared_Params

import os
import sys
import re
import math
from math import atan2
from fractions import Fraction


# --------------------------------------------------
# Shared params
# --------------------------------------------------
Shared_Params()


# --------------------------------------------------
# Basic environment
# --------------------------------------------------
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
uiapp = __revit__
active_view = doc.ActiveView

script_dir, script_name = os.path.split(__file__)

FAMILY_NAME = 'WS'
FAMILY_TYPE = 'WS'
FAMILY_FILE = 'WS.rfa'
family_path = os.path.join(script_dir, FAMILY_FILE)

RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)


# --------------------------------------------------
# General helpers
# --------------------------------------------------
def get_id_value(id_obj):
    if id_obj is None:
        return None
    try:
        if RevitINT > 2025:
            return id_obj.Value
        else:
            return id_obj.IntegerValue
    except:
        try:
            return id_obj.Value
        except:
            return id_obj.IntegerValue


def safe_taskdialog(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        print("{}: {}".format(title, message))


def get_lookup_param(element, param_name):
    try:
        return element.LookupParameter(param_name)
    except:
        return None


def set_param_if_possible(element, param_name, value):
    try:
        p = element.LookupParameter(param_name)
        if not p or p.IsReadOnly:
            return False
        p.Set(value)
        return True
    except:
        return False


def round_up_to_nearest_quarter_foot(value_in_feet):
    try:
        value_in_inches = value_in_feet * 12.0
        rounded_inches = math.ceil(value_in_inches * 4.0) / 4.0
        return rounded_inches / 12.0
    except:
        return value_in_feet


def parse_size_to_feet(size_string):
    """
    Converts strings like:
      2"
      2 1/2"
      2-1/2"
      2.5"
    into feet.
    """
    if not size_string:
        raise Exception("Overall Size parameter is empty.")

    s = str(size_string).strip()

    def frac_repl(match):
        whole = match.group(1)
        frac = match.group(2)
        whole_val = int(whole) if whole else 0
        frac_val = float(Fraction(frac))
        return str(whole_val + frac_val)

    s = re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)', frac_repl, s)

    m = re.search(r'[-+]?\d*\.?\d+', s)
    if not m:
        raise Exception("Could not parse Overall Size value: '{}'".format(size_string))

    inches = float(m.group(0))
    return inches / 12.0


def get_element_connectors(element):
    connectors = []
    try:
        connectors = list(element.ConnectorManager.Connectors)
    except:
        pass

    if not connectors:
        try:
            connectors = list(element.MEPModel.ConnectorManager.Connectors)
        except:
            pass

    return connectors


def get_pipe_centerline(pipe):
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    raise Exception("Selected pipe does not have a valid centerline.")


def get_level_from_pipe(pipe):
    try:
        if pipe.LevelId and pipe.LevelId != DB.ElementId.InvalidElementId:
            return doc.GetElement(pipe.LevelId)
    except:
        pass
    return None


def validate_pipe(pipe):
    if not pipe:
        raise Exception("No pipe selected.")

    cat = getattr(pipe, "Category", None)
    if not cat or get_id_value(cat.Id) != int(DB.BuiltInCategory.OST_FabricationPipework):
        raise Exception("Selected element is not fabrication pipework.")

    _ = get_pipe_centerline(pipe)

    connectors = get_element_connectors(pipe)
    if len(connectors) < 2:
        raise Exception("Selected fabrication pipe does not have at least two connectors.")

    overall_size = get_parameter_value_by_name_AsString(pipe, 'Overall Size')
    if not overall_size:
        raise Exception("Selected pipe does not have a valid 'Overall Size' value.")

    return True


def get_pipe_direction(pipe):
    curve = get_pipe_centerline(pipe)
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    direction = (p1 - p0)
    if direction.GetLength() == 0:
        raise Exception("Pipe direction is invalid.")

    return direction.Normalize()


def get_horizontal_pipe_direction(pipe):
    curve = get_pipe_centerline(pipe)
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    v = DB.XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)

    if v.GetLength() < 1e-8:
        raise Exception("Pipe is vertical or too close to vertical for wall sleeve placement.")

    return v.Normalize()


# --------------------------------------------------
# Linked wall helpers
# --------------------------------------------------
def get_wall_thickness_curve_and_plan_normal(document, link_ref):
    link_instance = document.GetElement(link_ref.ElementId)
    if not link_instance or not isinstance(link_instance, DB.RevitLinkInstance):
        raise Exception("Invalid linked wall selection.")

    link_doc = link_instance.GetLinkDocument()
    if not link_doc:
        raise Exception("Could not access linked document.")

    wall = link_doc.GetElement(link_ref.LinkedElementId)
    if not wall or not isinstance(wall, DB.Wall):
        raise Exception("Selected linked element is not a wall.")

    wall_type = wall.WallType
    if not wall_type:
        raise Exception("Selected wall does not have a valid wall type.")

    width_param = wall_type.get_Parameter(DB.BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
    if not width_param:
        raise Exception("Could not get wall thickness.")

    thickness = width_param.AsDouble()

    location = wall.Location
    if not isinstance(location, LocationCurve):
        raise Exception("Selected wall does not have a valid location curve.")

    try:
        link_transform = link_instance.GetTotalTransform()
    except:
        link_transform = link_instance.GetTransform()

    wall_curve_host = location.Curve.CreateTransformed(link_transform)

    try:
        wall_normal = link_transform.OfVector(wall.Orientation)
        wall_plan_normal = DB.XYZ(wall_normal.X, wall_normal.Y, 0.0)
    except:
        w0 = wall_curve_host.GetEndPoint(0)
        w1 = wall_curve_host.GetEndPoint(1)
        wall_dir = DB.XYZ(w1.X - w0.X, w1.Y - w0.Y, 0.0)
        wall_plan_normal = DB.XYZ(-wall_dir.Y, wall_dir.X, 0.0)

    if wall_plan_normal.GetLength() < 1e-8:
        raise Exception("Could not determine wall plan normal.")

    return thickness, wall_curve_host, wall_plan_normal.Normalize()


def get_sloped_pipe_wall_intersection_point(pipe_curve, wall_curve, wall_plan_normal):
    """
    Solve wall crossing in plan, then rebuild the actual 3D point on the pipe
    using the same t parameter.
    """
    p0 = pipe_curve.GetEndPoint(0)
    p1 = pipe_curve.GetEndPoint(1)
    wall_origin = wall_curve.GetEndPoint(0)

    p0_xy = DB.XYZ(p0.X, p0.Y, 0.0)
    p1_xy = DB.XYZ(p1.X, p1.Y, 0.0)
    o_xy = DB.XYZ(wall_origin.X, wall_origin.Y, 0.0)
    n_xy = DB.XYZ(wall_plan_normal.X, wall_plan_normal.Y, 0.0)

    if n_xy.GetLength() < 1e-8:
        return None

    n_xy = n_xy.Normalize()

    d0 = n_xy.DotProduct(p0_xy - o_xy)
    d1 = n_xy.DotProduct(p1_xy - o_xy)

    if abs(d0) < 1e-9:
        return p0
    if abs(d1) < 1e-9:
        return p1

    if abs(d0 - d1) < 1e-9:
        return None

    t = d0 / float(d0 - d1)

    if t < -1e-9 or t > 1.0 + 1e-9:
        return None

    t = max(0.0, min(1.0, t))
    return p0 + ((p1 - p0) * t)


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
class WSFamilyManager(object):
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
            safe_taskdialog("Error", "Family file not found:\n{}".format(self.family_path))
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
            safe_taskdialog("Error", "Family load error:\n{}".format(str(e)))
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
            safe_taskdialog("Error", "Family symbol activation error:\n{}".format(str(e)))
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
            safe_taskdialog(
                "Error",
                "Type '{}' not found in family '{}'.".format(type_name, self.family_name)
            )
            return None

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


# --------------------------------------------------
# Family symbol ready
# --------------------------------------------------
family_manager = WSFamilyManager(doc, FAMILY_NAME, family_path)
famsymb = family_manager.get_ready_symbol(FAMILY_TYPE)

if not famsymb:
    sys.exit(1)


# --------------------------------------------------
# Selection filters
# --------------------------------------------------
class LinkedWallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.RevitLinkInstance)

    def AllowReference(self, reference, position):
        try:
            link_instance = doc.GetElement(reference.ElementId)
            if not isinstance(link_instance, DB.RevitLinkInstance):
                return False

            linked_doc = link_instance.GetLinkDocument()
            if not linked_doc:
                return False

            linked_elem = linked_doc.GetElement(reference.LinkedElementId)
            return isinstance(linked_elem, DB.Wall)
        except:
            return False


class PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            return element.Category and get_id_value(element.Category.Id) == int(DB.BuiltInCategory.OST_FabricationPipework)
        except:
            return False

    def AllowReference(self, reference, position):
        return False


# --------------------------------------------------
# Selection
# --------------------------------------------------
def select_fabrication_pipes():
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        PipeSelectionFilter(),
        "Select one or more MEP Fabrication Pipes (press Esc to cancel)"
    )

    pipes = []
    for ref in refs:
        pipe = doc.GetElement(ref.ElementId)
        validate_pipe(pipe)
        pipes.append(pipe)

    if not pipes:
        raise Exception("No valid fabrication pipes were selected.")

    return pipes


def select_linked_wall():
    ref = uidoc.Selection.PickObject(
        ObjectType.LinkedElement,
        LinkedWallSelectionFilter(),
        "Select a wall in a Revit link"
    )
    if not ref:
        raise Exception("No linked wall was selected.")
    return ref


# --------------------------------------------------
# Placement
# --------------------------------------------------
def create_family_instance_at_point(insertion_point, symbol, level):
    try:
        if level:
            return doc.Create.NewFamilyInstance(
                insertion_point,
                symbol,
                level,
                DB.Structure.StructuralType.NonStructural
            )
    except:
        pass

    return doc.Create.NewFamilyInstance(
        insertion_point,
        symbol,
        DB.Structure.StructuralType.NonStructural
    )


def populate_instance_parameters(new_family_instance, pipe, wall_thickness, overall_size_feet, level):
    sleeve_diameter = overall_size_feet + (1.0 / 12.0)
    sleeve_diameter = round_up_to_nearest_quarter_foot(sleeve_diameter)

    set_param_if_possible(new_family_instance, 'Diameter', sleeve_diameter)
    set_param_if_possible(new_family_instance, 'Length', wall_thickness)

    params = {
        'FP_Product Entry': 'Overall Size',
        'FP_Service Name': 'Fabrication Service Name',
        'FP_Service Abbreviation': 'Fabrication Service Abbreviation'
    }

    for fam_param, pipe_param in params.items():
        value = get_parameter_value_by_name_AsString(pipe, pipe_param)
        if value is None:
            value = ''
        try:
            set_parameter_by_name(new_family_instance, fam_param, value)
        except:
            set_param_if_possible(new_family_instance, fam_param, value)

    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly and level:
        try:
            schedule_level_param.Set(level.Id)
        except:
            pass


def rotate_instance_to_pipe_direction(new_family_instance, insertion_point, pipe_direction):
    angle = atan2(pipe_direction.Y, pipe_direction.X)
    axis = DB.Line.CreateBound(
        insertion_point,
        DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1.0)
    )
    DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)


def place_and_modify_family(pipe, wall_ref, famsymb):
    centerline_curve = get_pipe_centerline(pipe)

    wall_thickness, wall_curve, wall_plan_normal = \
        get_wall_thickness_curve_and_plan_normal(doc, wall_ref)

    flat_pipe_direction = get_horizontal_pipe_direction(pipe)

    intersection_point = get_sloped_pipe_wall_intersection_point(
        centerline_curve,
        wall_curve,
        wall_plan_normal
    )

    if not intersection_point:
        raise Exception(
            "Pipe Id {} does not cross the selected linked wall in plan."
            .format(get_id_value(pipe.Id))
        )

    level = get_level_from_pipe(pipe)

    insertion_point = intersection_point - flat_pipe_direction * (wall_thickness / 2.0)

    new_family_instance = create_family_instance_at_point(insertion_point, famsymb, level)
    if not new_family_instance:
        raise Exception("Failed to create family instance for pipe Id {}.".format(get_id_value(pipe.Id)))

    overall_size = get_parameter_value_by_name_AsString(pipe, 'Overall Size')
    overall_size_feet = parse_size_to_feet(overall_size)

    populate_instance_parameters(
        new_family_instance,
        pipe,
        wall_thickness,
        overall_size_feet,
        level
    )

    rotate_instance_to_pipe_direction(new_family_instance, insertion_point, flat_pipe_direction)

    return new_family_instance


# --------------------------------------------------
# Main
# --------------------------------------------------
def main():
    try:
        pipes = select_fabrication_pipes()
        wall_ref = select_linked_wall()

    except OperationCanceledException:
        return
    except Exception as ex:
        safe_taskdialog("Error", str(ex))
        return

    placed_count = 0
    failed = []

    t = None
    try:
        t = Transaction(doc, 'Place Trimble Wall Sleeve Family')
        t.Start()

        for pipe in pipes:
            try:
                place_and_modify_family(pipe, wall_ref, famsymb)
                placed_count += 1
            except Exception as pipe_ex:
                failed.append(
                    "Pipe Id {}: {}".format(get_id_value(pipe.Id), str(pipe_ex))
                )

        if placed_count == 0 and failed:
            t.RollBack()
            safe_taskdialog(
                "Wall Sleeve Placement",
                "No sleeves were placed.\n\n{}".format("\n".join(failed[:20]))
            )
            return

        t.Commit()

    except OperationCanceledException:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        return

    except Exception as ex:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        safe_taskdialog("Error", "Operation failed:\n{}".format(str(ex)))
        return

    finally:
        if t:
            t.Dispose()

    if failed:
        msg = "Placed {} sleeve(s).\nFailed {} pipe(s).".format(placed_count, len(failed))
        msg += "\n\n" + "\n".join(failed[:20])
        if len(failed) > 20:
            msg += "\n..."
        safe_taskdialog("Wall Sleeve Placement", msg)
    else:
        safe_taskdialog("Wall Sleeve Placement", "Placed {} sleeve(s).".format(placed_count))


if __name__ == '__main__':
    main()