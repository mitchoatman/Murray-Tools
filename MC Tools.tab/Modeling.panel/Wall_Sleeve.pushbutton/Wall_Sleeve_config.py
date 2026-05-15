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
    BuiltInCategory,
    FamilySymbol,
    LocationCurve,
    Transaction,
    LinkElementId,
    Wall
)
from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsInteger
)
import re
import os
from math import atan2, degrees
from fractions import Fraction
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

# Get file path information
path, filename = os.path.split(__file__)
family_pathCC = os.path.join(path, 'Round Wall Sleeve.rfa')

# Get Revit application and document objects
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
active_view = doc.ActiveView

FamilyName = 'Round Wall Sleeve'
FamilyType = 'Round Wall Sleeve'

# Define file path for sleeve length storage
folder_name = r"c:\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Sleeve.txt')

# Create directory and default file if they don't exist
if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('6')  # Default sleeve length of 6

# Read sleeve length from file
with open(filepath, 'r') as f:
    SleeveLength = float(f.read())

# Get the level associated with the active view
try:
    level = active_view.GenLevel
except:
    level = None


def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        pass


class LinkedWallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.RevitLinkInstance)

    def AllowReference(self, reference, position):
        try:
            link_instance = doc.GetElement(reference.ElementId)
            if not isinstance(link_instance, DB.RevitLinkInstance):
                return False
            linked_doc = link_instance.GetLinkDocument()
            linked_elem = linked_doc.GetElement(reference.LinkedElementId)
            return isinstance(linked_elem, DB.Wall)
        except:
            return False


# --------------------------------------------------
# Robust family loader
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


class RoundWallSleeveFamilyManager(object):
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


family_manager = RoundWallSleeveFamilyManager(doc, FamilyName, family_pathCC)
famsymb = family_manager.get_ready_symbol(FamilyType)

if not famsymb:
    raise Exception("Family symbol '{}' was not found or could not be activated.".format(FamilyType))


# Optimized diameter mapping using a dictionary
DIAMETER_MAP = {
    (0.0, 1.0): 2.0,
    (1.0, 1.25): 2.5,
    (1.25, 1.5): 3.0,
    (1.5, 2.5): 4.0,
    (2.5, 3.5): 5.0,
    (3.5, 4.5): 6.0,
    (4.5, 7.5): 8.0,
    (7.5, 8.5): 10.0,
    (8.5, 10.5): 12.0,
    (10.5, 14.5): 16.0,
    (14.5, 16.5): 18.0,
    (16.5, 18.5): 20.0
}


# Custom selection filter for pipes
class PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            return element.Category and element.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationPipework)
        except:
            return False

    def AllowReference(self, reference, position):
        return False


def select_fabrication_pipes():
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        PipeSelectionFilter(),
        "Select one or more MEP Fabrication Pipes (press Esc to finish)"
    )
    pipes = [doc.GetElement(ref.ElementId) for ref in refs]
    if not pipes:
        raise Exception("No pipes selected.")
    return pipes


def select_linked_wall():
    ref = uidoc.Selection.PickObject(
        ObjectType.LinkedElement,
        LinkedWallSelectionFilter(),
        "Select a wall in a Revit link"
    )
    return ref


def get_wall_thickness_and_location(doc, link_ref):
    link_id = link_ref.LinkedElementId
    link_doc = doc.GetElement(link_ref.ElementId).GetLinkDocument()
    if not link_doc:
        raise Exception("Could not access linked document.")

    wall = link_doc.GetElement(link_id)
    if not wall:
        raise Exception("Selected element is not a valid wall.")

    thickness = wall.WallType.get_Parameter(DB.BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
    location = wall.Location
    if isinstance(location, LocationCurve):
        return thickness, location.Curve

    raise Exception("Selected wall does not have a valid location curve.")


def get_pipe_centerline(pipe):
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    raise Exception("Selected pipe does not have a valid centerline.")


def get_diameter_from_size(pipe_diameter):
    pipe_diameter *= 12
    for (min_val, max_val), sleeve_size in DIAMETER_MAP.items():
        if min_val < pipe_diameter < max_val:
            return sleeve_size / 12
    return 2.0 / 12


def project_wall_curve_to_pipe_plane(wall_curve, pipe_curve):
    pipe_start = pipe_curve.GetEndPoint(0)
    pipe_end = pipe_curve.GetEndPoint(1)
    pipe_direction = (pipe_end - pipe_start).Normalize()
    plane_normal = DB.XYZ(0, 0, 1)
    if abs(pipe_direction.Z) > 0.99:
        plane_normal = DB.XYZ(1, 0, 0)
    plane = DB.Plane.CreateByNormalAndOrigin(plane_normal, pipe_start)

    wall_start = wall_curve.GetEndPoint(0)
    wall_end = wall_curve.GetEndPoint(1)

    uv_start, distance_start = plane.Project(wall_start)
    uv_end, distance_end = plane.Project(wall_end)

    origin = plane.Origin
    x_axis = plane.XVec
    y_axis = plane.YVec
    projected_start = origin + x_axis * uv_start.U + y_axis * uv_start.V
    projected_end = origin + x_axis * uv_end.U + y_axis * uv_end.V

    try:
        projected_curve = DB.Line.CreateBound(projected_start, projected_end)
        return projected_curve
    except Exception as e:
        raise Exception("Error projecting wall curve: {}".format(str(e)))


def place_and_modify_family(pipe, wall_ref, famsymb):
    centerline_curve = get_pipe_centerline(pipe)
    wall_thickness, wall_curve = get_wall_thickness_and_location(doc, wall_ref)

    wall_curve = project_wall_curve_to_pipe_plane(wall_curve, centerline_curve)

    intersection_result = centerline_curve.Intersect(wall_curve)
    if intersection_result != DB.SetComparisonResult.Overlap:
        raise Exception("Pipe does not intersect with the projected wall curve.")

    result_array = clr.Reference[DB.IntersectionResultArray]()
    centerline_curve.Intersect(wall_curve, result_array)
    if result_array.Value and result_array.Value.Size > 0:
        intersection_point = result_array.Value[0].XYZPoint
    else:
        raise Exception("Failed to retrieve intersection point.")

    connectors = list(pipe.ConnectorManager.Connectors)
    if len(connectors) < 2:
        raise Exception("Selected pipe does not have enough connectors.")

    conn1, conn2 = connectors[0], connectors[1]
    nearest_conn = min([conn1, conn2], key=lambda c: intersection_point.DistanceTo(c.Origin))
    other_conn = conn2 if nearest_conn == conn1 else conn1
    pipe_direction = (other_conn.Origin - nearest_conn.Origin).Normalize()
    offset_distance = wall_thickness / 2.0
    insertion_point = intersection_point - pipe_direction * offset_distance

    try:
        if level:
            new_family_instance = doc.Create.NewFamilyInstance(
                insertion_point,
                famsymb,
                level,
                DB.Structure.StructuralType.NonStructural
            )
        else:
            new_family_instance = doc.Create.NewFamilyInstance(
                insertion_point,
                famsymb,
                DB.Structure.StructuralType.NonStructural
            )
    except:
        new_family_instance = doc.Create.NewFamilyInstance(
            insertion_point,
            famsymb,
            DB.Structure.StructuralType.NonStructural
        )

    if not new_family_instance:
        raise Exception("Failed to create family instance.")

    overall_size = get_parameter_value_by_name_AsString(pipe, 'Overall Size')
    cleaned_size = re.sub(r'["]', '', overall_size.strip())

    try:
        diameter = float(cleaned_size)
    except ValueError:
        match = re.match(r'(?:(\d+)[-\s])?(\d+/\d+)', cleaned_size)
        if match:
            integer_part, fraction_part = match.groups()
            diameter = float(Fraction(fraction_part))
            if integer_part:
                diameter += float(integer_part)
        else:
            diameter = 0.5

    diameter = diameter / 12
    sleeve_diameter = get_diameter_from_size(diameter)
    set_parameter_by_name(new_family_instance, 'Diameter', sleeve_diameter)
    set_parameter_by_name(new_family_instance, 'Length', wall_thickness)

    vec = other_conn.Origin - nearest_conn.Origin
    angle = atan2(vec.Y, vec.X)
    axis = DB.Line.CreateBound(
        insertion_point,
        DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1)
    )
    DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)

    params = {
        'FP_Product Entry': 'Overall Size',
        'FP_Service Name': 'Fabrication Service Name',
        'FP_Service Abbreviation': 'Fabrication Service Abbreviation'
    }
    for fam_param, pipe_param in params.items():
        value = get_parameter_value_by_name_AsString(pipe, pipe_param)
        set_parameter_by_name(new_family_instance, fam_param, value)

    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly and level:
        schedule_level_param.Set(level.Id)


# Main execution loop
failed = []
t = None

try:
    pipes = select_fabrication_pipes()
    wall_ref = select_linked_wall()

    t = Transaction(doc, 'Place Wall Sleeve Family')
    t.Start()

    for pipe in pipes:
        try:
            place_and_modify_family(pipe, wall_ref, famsymb)
        except Exception as pipe_error:
            try:
                failed.append("Pipe {}: {}".format(pipe.Id.IntegerValue, str(pipe_error)))
            except:
                failed.append(str(pipe_error))

    if failed and len(failed) == len(pipes):
        t.RollBack()
        show_message("Wall Sleeve Placement", "No sleeves were placed.\n\n{}".format("\n".join(failed[:20])))
    else:
        t.Commit()
        if failed:
            show_message(
                "Wall Sleeve Placement",
                "Placement completed with some issues.\n\n{}".format("\n".join(failed[:20]))
            )

except Exception as e:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    show_message("Error", "Error during execution:\n{}".format(str(e)))

finally:
    if t:
        t.Dispose()