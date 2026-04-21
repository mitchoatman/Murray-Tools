from Autodesk.Revit import DB
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    BuiltInCategory,
    FamilySymbol,
    LocationCurve,
    Transaction,
    LinkElementId,
    Wall,
    RevitLinkInstance
)
from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsInteger
)
from Autodesk.Revit.UI import TaskDialog
import re
from math import atan2, degrees
from fractions import Fraction
import os
from Parameters.Add_SharedParameters import Shared_Params
import clr
import sys

Shared_Params()

path, filename = os.path.split(__file__)
NewFilename = r'\WS.rfa'
family_pathCC = path + NewFilename

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

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
        t = Transaction(doc, 'Load Trimble Wall Sleeve Family')
        t.Start()
        families = FilteredElementCollector(doc).OfClass(Family)
        FamilyName = 'WS'
        if not any(f.Name == FamilyName for f in families):
            load_options = FamilyLoadOptions()
            loaded_family = clr.StrongBox[DB.Family]()
            success = doc.LoadFamily(family_path, load_options, loaded_family)
            if not success or not loaded_family.Value:
                raise Exception("Failed to load family")
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



try:
    load_family(family_pathCC)
except Exception as e:
    print("Family load failed: {}".format(e))
    sys.exit(1)

collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
FamilyName = 'WS'
FamilyType = 'WS'
famsymb = None
for fs in collector:
    if fs.Family.Name == FamilyName and fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType:
        famsymb = fs
        break

if famsymb:
    t = None
    try:
        t = Transaction(doc, 'Activate Family Symbol')
        t.Start()
        famsymb.Activate()
        doc.Regenerate()
        t.Commit()
    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        print("Family symbol activation error: {}".format(e))
        sys.exit(1)
    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()

# Custom selection filter for walls in linked model
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

# Custom selection filter for pipes
class PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            return element.Category and get_id_value(element.Category.Id) == int(BuiltInCategory.OST_FabricationPipework)
        except:
            return False

    def AllowReference(self, reference, position):
        return False

def select_fabrication_pipes():
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, PipeSelectionFilter(), 
                                          "Select one or more MEP Fabrication Pipes (press Esc to finish)")
        pipes = [doc.GetElement(ref.ElementId) for ref in refs]
        if not pipes:
            print("Error: No pipes selected.")
            raise Exception("No pipes selected.")
        return pipes
    except Exception as e:
        print("Error selecting pipes: {}".format(str(e)))
        raise

def select_linked_wall():
    try:
        ref = uidoc.Selection.PickObject(ObjectType.LinkedElement, LinkedWallSelectionFilter(), 
                                        "Select a wall in a Revit link")
        return ref
    except Exception as e:
        print("Error selecting linked wall: {}".format(str(e)))
        raise

def get_wall_thickness_and_location(doc, link_ref):
    link_id = link_ref.LinkedElementId
    link_doc = doc.GetElement(link_ref.ElementId).GetLinkDocument()
    if not link_doc:
        print("Error: Could not access linked document.")
        raise Exception("Could not access linked document.")
    wall = link_doc.GetElement(link_id)
    if not wall:
        print("Error: Selected element is not a valid wall.")
        raise Exception("Selected element is not a valid wall.")
    thickness = wall.WallType.get_Parameter(DB.BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
    location = wall.Location
    if isinstance(location, LocationCurve):
        return thickness, location.Curve
    print("Error: Selected wall does not have a valid location curve.")
    raise Exception("Selected wall does not have a valid location curve.")

def get_pipe_centerline(pipe):
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    print("Error: Selected pipe does not have a valid centerline.")
    raise Exception("Selected pipe does not have a valid centerline.")

def project_wall_curve_to_pipe_plane(wall_curve, pipe_curve):
    pipe_start = pipe_curve.GetEndPoint(0)
    pipe_end = pipe_curve.GetEndPoint(1)
    pipe_direction = (pipe_end - pipe_start).Normalize()
    plane_normal = DB.XYZ(0, 0, 1)  # Assuming pipe is horizontal, use Z-axis as normal
    if abs(pipe_direction.Z) > 0.99:  # If pipe is vertical, use X-axis as normal
        plane_normal = DB.XYZ(1, 0, 0)
    plane = DB.Plane.CreateByNormalAndOrigin(plane_normal, pipe_start)
    
    # Project wall curve endpoints
    wall_start = wall_curve.GetEndPoint(0)
    wall_end = wall_curve.GetEndPoint(1)
    
    # Project points onto the plane
    uv_start, distance_start = plane.Project(wall_start)
    uv_end, distance_end = plane.Project(wall_end)
    
    # Convert UV coordinates to XYZ using plane's basis vectors
    origin = plane.Origin
    x_axis = plane.XVec
    y_axis = plane.YVec
    projected_start = origin + x_axis * uv_start.U + y_axis * uv_start.V
    projected_end = origin + x_axis * uv_end.U + y_axis * uv_end.V
    
    try:
        projected_curve = DB.Line.CreateBound(projected_start, projected_end)
        return projected_curve
    except Exception as e:
        print("Error projecting wall curve: {}".format(str(e)))
        raise

def get_first_pipe_direction(pipes):
    if not pipes:
        print("Error: No pipes to determine direction.")
        raise Exception("No pipes to determine direction.")
    first_pipe = pipes[0]
    connectors = list(first_pipe.ConnectorManager.Connectors)
    conn1, conn2 = connectors[0], connectors[1]
    return (conn2.Origin - conn1.Origin).Normalize()

def place_and_modify_family(pipe, wall_ref, famsymb, fixed_pipe_direction):
    try:
        centerline_curve = get_pipe_centerline(pipe)
        wall_thickness, wall_curve = get_wall_thickness_and_location(doc, wall_ref)
        
        # Project wall curve to pipe's plane
        wall_curve = project_wall_curve_to_pipe_plane(wall_curve, centerline_curve)
        
        # Get intersection point between pipe and projected wall curve
        intersection_result = centerline_curve.Intersect(wall_curve)
        if intersection_result != DB.SetComparisonResult.Overlap:
            print("Error: Pipe does not intersect with the projected wall curve.")
            raise Exception("Pipe does not intersect with the projected wall curve.")
        
        # Get intersection point
        result_array = clr.Reference[DB.IntersectionResultArray]()
        centerline_curve.Intersect(wall_curve, result_array)
        if result_array.Value and result_array.Value.Size > 0:
            intersection_point = result_array.Value[0].XYZPoint
        else:
            print("Error: Failed to retrieve intersection point.")
            raise Exception("Failed to retrieve intersection point.")

        # Get reference level from pipe
        level = doc.GetElement(pipe.LevelId)
        
        # Calculate offset to align sleeve start with wall face
        offset_distance = wall_thickness / 2.0
        insertion_point = intersection_point - fixed_pipe_direction * offset_distance

        # Create new family instance at adjusted point
        new_family_instance = doc.Create.NewFamilyInstance(insertion_point, famsymb, 
                                                         level, DB.Structure.StructuralType.NonStructural)
        if not new_family_instance:
            print("Error: Failed to create family instance.")
            raise Exception("Failed to create family instance.")

        # Calculate and set diameter
        def frac2string(s):
            i, f = s.groups(0)
            f = Fraction(f)
            return str(int(i) + float(f))

        overall_size = get_parameter_value_by_name_AsString(pipe, 'Overall Size')
        if '/' in overall_size:
            diameter = float(re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)[^\d.]', frac2string, overall_size)) / 12 + 0.0833333
        else:
            diameter = float(re.sub(r'[^\d.]', '', overall_size)) / 12 + 0.0833333
        set_parameter_by_name(new_family_instance, 'Diameter', diameter)
        set_parameter_by_name(new_family_instance, 'Length', wall_thickness)

        # Align family with fixed pipe direction
        angle = atan2(fixed_pipe_direction.Y, fixed_pipe_direction.X)
        axis = DB.Line.CreateBound(insertion_point, 
                                 DB.XYZ(insertion_point.X, insertion_point.Y, intersection_point.Z + 1))
        DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)

        # Set family parameters
        params = {
            'FP_Product Entry': 'Overall Size',
            'FP_Service Name': 'Fabrication Service Name',
            'FP_Service Abbreviation': 'Fabrication Service Abbreviation'
        }
        for fam_param, pipe_param in params.items():
            value = get_parameter_value_by_name_AsString(pipe, pipe_param)
            set_parameter_by_name(new_family_instance, fam_param, value)
        
        schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
        schedule_level_param.Set(level.Id)
        return True
    except Exception as e:
        print("Family placement error: {}".format(str(e)))
        return False

from Autodesk.Revit.Exceptions import OperationCanceledException

try:
    with Transaction(doc, 'Place Trimble Wall Sleeve Family') as t:
        t.Start()
        pipes = select_fabrication_pipes()
        wall_ref = select_linked_wall()
        fixed_pipe_direction = get_first_pipe_direction(pipes)
        for pipe in pipes:
            if not place_and_modify_family(pipe, wall_ref, famsymb, fixed_pipe_direction):
                print("Error: Failed to place sleeve for a pipe.")
                continue
        t.Commit()
except OperationCanceledException:
    pass
except Exception as e:
    print("Error during operation: {}".format(str(e)))