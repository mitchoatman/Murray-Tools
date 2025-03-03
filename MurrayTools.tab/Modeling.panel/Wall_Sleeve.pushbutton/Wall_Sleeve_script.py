from Autodesk.Revit import DB
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    BuiltInCategory,
    FamilySymbol,
    LocationCurve,
    Transaction
    )
from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsInteger
    )
import re
from math import atan2, degrees
from fractions import Fraction
import os

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

def select_fabrication_pipe():
    """Prompt user to select an MEP Fabrication Pipe"""
    return doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, 
                                                    "Select an MEP Fabrication Pipe").ElementId)

def pick_point():
    """Prompt user to pick a point along pipe centerline"""
    return uidoc.Selection.PickPoint("Pick a point along the centerline of the pipe")

def get_pipe_centerline(pipe):
    """Get the centerline curve from a pipe element"""
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    raise Exception("The selected element does not have a valid centerline.")

def project_point_on_curve(point, curve):
    """Project a point onto a curve and return the XYZ point"""
    return curve.Project(point).XYZPoint

def get_diameter_from_size(pipe_diameter):
    """Convert pipe diameter to sleeve diameter using mapping"""
    pipe_diameter *= 12  # Convert to inches
    for (min_val, max_val), sleeve_size in DIAMETER_MAP.items():
        if min_val < pipe_diameter < max_val:
            return sleeve_size / 12  # Convert back to feet
    return 2.0 / 12  # Default minimum size

def place_and_modify_family(pipe, famsymb):
    """Place and configure a family instance aligned with a pipe"""
    centerline_curve = get_pipe_centerline(pipe)
    picked_point = pick_point()
    projected_point = project_point_on_curve(picked_point, centerline_curve)
    insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)
    
    # Create new family instance
    new_family_instance = doc.Create.NewFamilyInstance(insertion_point, famsymb, 
                                                     DB.Structure.StructuralType.NonStructural)

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
            # Default to 0.5 if parsing fails
            print("Warning: Could not parse diameter '{0}', defaulting to 0.5\"".format(overall_size))
            diameter = 0.5
    
    diameter = diameter / 12  # Convert inches to feet
    set_parameter_by_name(new_family_instance, 'Diameter', get_diameter_from_size(diameter))
    set_parameter_by_name(new_family_instance, 'Length', SleeveLength)
    
    # Align family with pipe
    connectors = list(pipe.ConnectorManager.Connectors)
    conn1, conn2 = connectors[0], connectors[1]
    nearest_conn = min([conn1, conn2], key=lambda c: picked_point.DistanceTo(c.Origin))
    other_conn = conn2 if nearest_conn == conn1 else conn1
    
    # Calculate rotation
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
        set_parameter_by_name(new_family_instance, fam_param, 
                           get_parameter_value_by_name_AsString(pipe, pipe_param))
    new_family_instance.LookupParameter("Schedule Level").Set(level.Id)

# Main execution loop
while True:
    try:
        with Transaction(doc, 'Place Wall Sleeve Family') as t:
            t.Start()
            pipe = select_fabrication_pipe()
            place_and_modify_family(pipe, famsymb)
            t.Commit()
    except:
        break