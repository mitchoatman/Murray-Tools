from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, BuiltInCategory, FamilySymbol, Transaction, XYZ
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsDouble
from math import atan2
import os
from Parameters.Add_SharedParameters import Shared_Params
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Selection filter using category IDs (same style as example script)
class PickByCategorySelectionFilter(ISelectionFilter):
    """Selection filter implementation"""
    def __init__(self, category_ids):
        self.category_ids = category_ids

    def AllowElement(self, element):
        if element.Category and element.Category.Id in self.category_ids:
            return True
        return False

    def AllowReference(self, reference, point):
        return False

def select_fabrication_parts():
    try:
        category_ids = [
            DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_FabricationPipework).Id
        ]
        msfilter = PickByCategorySelectionFilter(category_ids)
        selection = uidoc.Selection.PickObjects(ObjectType.Element, msfilter, "Select MEP Fabrication Pipework to Blockout")
        return [doc.GetElement(ref.ElementId) for ref in selection]
    except:
        return []

def get_combined_bounding_box(elements):
    if not elements:
        return None, 0, 0
    # Get directionality from first pipe
    first_pipe_bb = elements[0].get_BoundingBox(doc.ActiveView)
    if not first_pipe_bb:
        return None, 0, 0
    delta_x = abs(first_pipe_bb.Max.X - first_pipe_bb.Min.X)
    delta_y = abs(first_pipe_bb.Max.Y - first_pipe_bb.Min.Y)
    
    # Initialize with first pipe's bounding box
    combined_min = first_pipe_bb.Min
    combined_max = first_pipe_bb.Max
    
    # Adjust first pipe for insulation
    insulation_thickness = elements[0].InsulationThickness if hasattr(elements[0], 'InsulationThickness') and elements[0].HasInsulation else 0.0
    if delta_x > delta_y:  # X-direction
        combined_min = XYZ(combined_min.X, combined_min.Y - insulation_thickness, combined_min.Z - insulation_thickness)
        combined_max = XYZ(combined_max.X, combined_max.Y + insulation_thickness, combined_max.Z + insulation_thickness)
    else:  # Y-direction
        combined_min = XYZ(combined_min.X - insulation_thickness, combined_min.Y, combined_min.Z - insulation_thickness)
        combined_max = XYZ(combined_max.X + insulation_thickness, combined_max.Y, combined_max.Z + insulation_thickness)
    
    # Combine with other pipes
    for element in elements[1:]:
        bb = element.get_BoundingBox(doc.ActiveView)
        if not bb:
            continue
        insulation_thickness = element.InsulationThickness if hasattr(element, 'InsulationThickness') and element.HasInsulation else 0.0
        if delta_x > delta_y:  # X-direction
            element_min = XYZ(bb.Min.X, bb.Min.Y - insulation_thickness, bb.Min.Z - insulation_thickness)
            element_max = XYZ(bb.Max.X, bb.Max.Y + insulation_thickness, bb.Max.Z + insulation_thickness)
        else:  # Y-direction
            element_min = XYZ(bb.Min.X - insulation_thickness, bb.Min.Y, bb.Min.Z - insulation_thickness)
            element_max = XYZ(bb.Max.X + insulation_thickness, bb.Max.Y, bb.Max.Z + insulation_thickness)
        
        combined_min = XYZ(min(combined_min.X, element_min.X), min(combined_min.Y, element_min.Y), min(combined_min.Z, element_min.Z))
        combined_max = XYZ(max(combined_max.X, element_max.X), max(combined_max.Y, element_max.Y), max(combined_max.Z, element_max.Z))
    
    bbox = DB.BoundingBoxXYZ()
    bbox.Min = combined_min
    bbox.Max = combined_max
    return bbox, delta_x, delta_y

def get_duct_centerline(element):
    pipe_location = element.Location
    if isinstance(pipe_location, DB.LocationCurve):
        return pipe_location.Curve
    else:
        raise Exception("The selected element does not have a valid centerline.")

def get_center_of_bounding_box(bbox):
    if not bbox:
        return None
    center_x = (bbox.Min.X + bbox.Max.X) / 2
    center_y = (bbox.Min.Y + bbox.Max.Y) / 2
    center_z = (bbox.Min.Z + bbox.Max.Z) / 2
    return DB.XYZ(center_x, center_y, center_z)

def get_projected_point_on_axis(elements, user_point, bbox, delta_x, delta_y):
    centerline = get_duct_centerline(elements[0])
    start_point = centerline.GetEndPoint(0)
    end_point = centerline.GetEndPoint(1)
    direction = (end_point - start_point).Normalize()
    
    # Vector from start point to user-selected point
    vector_to_point = user_point - start_point
    
    # Project onto centerline
    projection_length = vector_to_point.DotProduct(direction)
    
    # Determine insertion point based on pipe direction
    center = get_center_of_bounding_box(bbox)
    if delta_x > delta_y:  # X-direction
        insertion_point = XYZ(start_point.X + direction.X * projection_length, center.Y, center.Z)
    else:  # Y-direction
        insertion_point = XYZ(center.X, start_point.Y + direction.Y * projection_length, center.Z)
    
    return insertion_point

def place_and_modify_family(elements, famsymb, annular_space):
    bbox, delta_x, delta_y = get_combined_bounding_box(elements)
    if not bbox:
        raise Exception("Unable to calculate combined bounding box")
    
    # Prompt user to select a point
    try:
        user_point = uidoc.Selection.PickPoint("Select location for blockout")
    except:
        raise Exception("Point selection cancelled or failed")
    
    # Project user-selected point onto centerline, aligned to center in other axes
    insertion_point = get_projected_point_on_axis(elements, user_point, bbox, delta_x, delta_y)
    
    # Get reference level from first pipe
    level = doc.GetElement(elements[0].LevelId)
    
    new_family_instance = doc.Create.NewFamilyInstance(insertion_point, famsymb, level, DB.Structure.StructuralType.NonStructural)
    
    # Calculate dimensions: Y for Width (X-direction), X for Width (Y-direction), Z for Height
    if delta_x > delta_y:  # X-direction
        width = (bbox.Max.Y - bbox.Min.Y) + annular_space / 12  # Y-dimension for width
    else:  # Y-direction
        width = (bbox.Max.X - bbox.Min.X) + annular_space / 12  # X-dimension for width
    height = (bbox.Max.Z - bbox.Min.Z) + annular_space / 12  # Z-dimension for height
    
    # Set parameters
    set_parameter_by_name(new_family_instance, 'Width', width)
    set_parameter_by_name(new_family_instance, 'Height', height)
    set_parameter_by_name(new_family_instance, 'Length', 0.25)
    
    # Get connectors for rotation
    centerline_curve = get_duct_centerline(elements[0])
    duct_connectors = list(elements[0].ConnectorManager.Connectors)
    connector1, connector2 = duct_connectors[0], duct_connectors[1]
    
    # Calculate rotation
    vec_x = connector2.Origin.X - connector1.Origin.X
    vec_y = connector2.Origin.Y - connector1.Origin.Y
    angle = atan2(vec_y, vec_x)
    axis = DB.Line.CreateBound(insertion_point, DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1))
    
    DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)
    
    # Set parameters
    set_parameter_by_name(new_family_instance, 'FP_Service Name', get_parameter_value_by_name_AsString(elements[0], 'Fabrication Service Name'))
    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param:
        schedule_level_param.Set(level.Id)
    else:
        print("Warning: 'Schedule Level' parameter not found in family instance")

# Family loading and setup
path, filename = os.path.split(__file__)
NewFilename = '\BLOCKOUT.rfa'
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
    print("Warning: Failed to read annular space, using default 1 inch")

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'BLOCKOUT'
FamilyType = 'BLOCKOUT'
Fam_is_in_project = any(f.Name == FamilyName for f in families)
family_pathCC = path + NewFilename

t = Transaction(doc, 'Load Trimble Wall Sleeve Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_pathCC, fload_handler)

collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
famsymb = next((fs for fs in collector if fs.Family.Name == FamilyName and fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType), None)

if famsymb:
    famsymb.Activate()
    doc.Regenerate()

t.Commit()

# Main execution
if famsymb:
    t = Transaction(doc, 'Place Trimble Wall Sleeve Family')
    t.Start()
    try:
        elements = select_fabrication_parts()
        if elements:
            place_and_modify_family(elements, famsymb, AnnularSpace)
        t.Commit()
    except Exception as e:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        print("Error during operation: {}".format(e))