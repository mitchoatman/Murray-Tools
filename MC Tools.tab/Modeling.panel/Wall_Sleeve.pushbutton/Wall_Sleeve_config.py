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
    Wall
)
from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsInteger
)
import re, clr, os
from math import atan2, degrees
from fractions import Fraction
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

# Get file path information
path, filename = os.path.split(__file__)
NewFilename = r'\Round Wall Sleeve.rfa'

# Get Revit application and document objects
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# Define file path for sleeve length storage
folder_name = "c:\Temp"
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
level = active_view.GenLevel

class LinkedWallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        # Allow all RevitLinkInstances to be clickable
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




# Define family loading options handler
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

# Check if family exists in project and load if necessary
families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'Round Wall Sleeve'
FamilyType = 'Round Wall Sleeve'
Fam_is_in_project = any(f.Name == FamilyName for f in families)
family_pathCC = path + NewFilename

# Load family if not present
t = Transaction(doc, 'Load Wall Sleeve Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_pathCC, fload_handler)

# Get family symbol
collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
famsymb = next((fs for fs in collector if fs.Family.Name == FamilyName and 
                fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType), None)

if famsymb:
    famsymb.Activate()
    doc.Regenerate()
else:
    print("Error: Family symbol 'Round Wall Sleeve' not found.")
t.Commit()

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
        return element.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationPipework)
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
        ref = uidoc.Selection.PickObject(ObjectType.LinkedElement, LinkedWallSelectionFilter(), "Select a wall in a Revit link")
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

def get_diameter_from_size(pipe_diameter):
    pipe_diameter *= 12  # Convert to inches
    for (min_val, max_val), sleeve_size in DIAMETER_MAP.items():
        if min_val < pipe_diameter < max_val:
            return sleeve_size / 12  # Convert back to feet
    print("Using default sleeve diameter: 2.0 inches")
    return 2.0 / 12  # Default minimum size

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

def place_and_modify_family(pipe, wall_ref, famsymb):
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

    # Calculate offset to align sleeve start with wall face
    connectors = list(pipe.ConnectorManager.Connectors)
    conn1, conn2 = connectors[0], connectors[1]
    nearest_conn = min([conn1, conn2], key=lambda c: intersection_point.DistanceTo(c.Origin))
    other_conn = conn2 if nearest_conn == conn1 else conn1
    pipe_direction = (other_conn.Origin - nearest_conn.Origin).Normalize()
    offset_distance = wall_thickness / 2.0
    insertion_point = intersection_point - pipe_direction * offset_distance

    # Create new family instance at adjusted point
    new_family_instance = doc.Create.NewFamilyInstance(insertion_point, famsymb, 
                                                     level, DB.Structure.StructuralType.NonStructural)
    if not new_family_instance:
        print("Error: Failed to create family instance.")
        raise Exception("Failed to create family instance.")

    # Calculate and set diameter with improved fraction handling
    overall_size = get_parameter_value_by_name_AsString(pipe, 'Overall Size')
    cleaned_size = re.sub(r'["]', '', overall_size.strip())  # Remove inch mark
    
    try:
        # Try direct float conversion for decimal or whole numbers
        diameter = float(cleaned_size)
    except ValueError:
        # Handle fractions like "1/2", "3/4", "5/8", "1/4"
        match = re.match(r'(?:(\d+)[-\s])?(\d+/\d+)', cleaned_size)
        if match:
            integer_part, fraction_part = match.groups()
            diameter = float(Fraction(fraction_part))
            if integer_part:
                diameter += float(integer_part)
        else:
            print("Warning: Could not parse diameter '{0}', defaulting to 0.5\"".format(overall_size))
            diameter = 0.5
    
    diameter = diameter / 12  # Convert inches to feet
    sleeve_diameter = get_diameter_from_size(diameter)
    set_parameter_by_name(new_family_instance, 'Diameter', sleeve_diameter)
    set_parameter_by_name(new_family_instance, 'Length', wall_thickness)
    
    # Align family with pipe
    vec = other_conn.Origin - nearest_conn.Origin
    angle = atan2(vec.Y, vec.X)
    axis = DB.Line.CreateBound(insertion_point, 
                             DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1))
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
    
    new_family_instance.LookupParameter("Schedule Level").Set(level.Id)

# Main execution loop
try:
    with Transaction(doc, 'Place Wall Sleeve Family') as t:
        t.Start()
        pipes = select_fabrication_pipes()
        wall_ref = select_linked_wall()
        for pipe in pipes:
            place_and_modify_family(pipe, wall_ref, famsymb)
        t.Commit()
except Exception as e:
    print("Error during execution: {}".format(str(e)))