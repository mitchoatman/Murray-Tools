from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    BuiltInCategory,
    FamilySymbol,
    Transaction,
    Line,
    XYZ,
    ViewType
)
from math import atan2
from fractions import Fraction
import re
import os
from SharedParam.Add_Parameters import Shared_Params
Shared_Params()

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString
)

# Get Revit document objects
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# File path setup
path, filename = os.path.split(__file__)
family_path = os.path.join(path, 'Round Floor Sleeve.rfa')

# Sleeve length storage
temp_folder = r"c:\Temp"
sleeve_length_file = os.path.join(temp_folder, 'Ribbon_Sleeve.txt')
if not os.path.exists(temp_folder):
    os.makedirs(temp_folder)
if not os.path.exists(sleeve_length_file):
    with open(sleeve_length_file, 'w') as f:
        f.write('6')
with open(sleeve_length_file, 'r') as f:
    sleeve_length = float(f.read())

# Diameter mapping for sleeve sizing (in inches, converted to feet)
DIAMETER_MAP = {
    (0.0, 1.0): 2.0, (1.0, 1.25): 2.5, (1.25, 1.5): 3.0,
    (1.5, 2.5): 4.0, (2.5, 3.5): 5.0, (3.5, 4.5): 6.0,
    (4.5, 7.5): 8.0, (7.5, 8.5): 10.0, (8.5, 10.5): 12.0,
    (10.5, 14.5): 16.0, (14.5, 16.5): 18.0, (16.5, 18.5): 20.0
}

def load_family():
    """Load 'Round Floor Sleeve' family if not present in project"""
    families = FilteredElementCollector(doc).OfClass(Family)
    family_name = 'Round Floor Sleeve'
    if not any(f.Name == family_name for f in families):
        t = Transaction(doc, 'Load Family')
        t.Start()
        doc.LoadFamily(family_path)
        t.Commit()
    
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
    return next((fs for fs in collector if fs.Family.Name == family_name and 
                fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == family_name), None)

def get_diameter_from_size(pipe_diameter):
    """Convert pipe diameter (in feet) to sleeve diameter (in feet)"""
    pipe_diameter *= 12  # Convert to inches
    for (min_val, max_val), sleeve_size in DIAMETER_MAP.items():
        if min_val < pipe_diameter <= max_val:
            return sleeve_size / 12  # Convert back to feet
    return 2.0 / 12  # Default minimum size

def get_pipe_centerline(pipe):
    """Get pipe centerline from its connectors"""
    connectors = list(pipe.ConnectorManager.Connectors)
    if len(connectors) >= 2:
        return Line.CreateBound(connectors[0].Origin, connectors[1].Origin)
    return None

def get_pipe_intersections(pipe, level):
    """Find intersection points between vertical pipe and level using manual calculation"""
    centerline = get_pipe_centerline(pipe)
    if not centerline:
        return []
    
    p0 = centerline.GetEndPoint(0)  # Start point
    p1 = centerline.GetEndPoint(1)  # End point
    normal = XYZ(0, 0, 1)
    plane_z = level.Elevation
    direction = (p1 - p0).Normalize()
    
    dot_product = direction.DotProduct(normal)
    if abs(dot_product) < 1e-6:  # Parallel, no intersection
        return []
    
    t = (plane_z - p0.Z) / direction.Z
    length = p0.DistanceTo(p1)
    if t < 0 or t > length:
        return []
    
    intersection_point = p0 + direction.Multiply(t)
    return [intersection_point]

def is_vertical_pipe(pipe):
    """Check if pipe is vertical using connectors"""
    connectors = list(pipe.ConnectorManager.Connectors)
    if len(connectors) < 2:
        return False
    direction = (connectors[1].Origin - connectors[0].Origin).Normalize()
    return abs(direction.Z) > 0.99  # Nearly vertical (cosine close to 1)

def is_duplicate_sleeve(intersection_point, existing_sleeves, tolerance=0.001):
    """Check if a sleeve already exists at the intersection point within tolerance"""
    for sleeve in existing_sleeves:
        sleeve_location = sleeve.Location.Point
        if (abs(sleeve_location.X - intersection_point.X) < tolerance and
            abs(sleeve_location.Y - intersection_point.Y) < tolerance and
            abs(sleeve_location.Z - intersection_point.Z) < tolerance):
            return True
    return False

def place_sleeve_at_intersection(pipe, intersection_point, family_symbol, level, existing_sleeves):
    """Place and configure sleeve family instance at intersection if no duplicate exists"""
    if is_duplicate_sleeve(intersection_point, existing_sleeves):
        return None
    
    new_instance = doc.Create.NewFamilyInstance(
        intersection_point, 
        family_symbol,
        level,
        DB.Structure.StructuralType.NonStructural
    )
    
    # Set diameter with improved fraction handling
    overall_size = pipe.get_Parameter(DB.BuiltInParameter.RBS_REFERENCE_OVERALLSIZE).AsString()
    cleaned_size = re.sub(r'["]', '', overall_size.strip())  # Remove " specifically
    
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
            print("Warning: Could not parse diameter '{0}', defaulting to 0.5\"".format(overall_size))
            diameter = 0.5
    
    diameter = diameter / 12  # Convert inches to feet
    new_instance.LookupParameter('Diameter').Set(get_diameter_from_size(diameter))
    
    # Set length and level
    new_instance.LookupParameter('Length').Set(sleeve_length)
    new_instance.LookupParameter('Schedule Level').Set(level.Id)
    
    # Align sleeve with pipe direction
    connectors = list(pipe.ConnectorManager.Connectors)
    vec = (connectors[1].Origin - connectors[0].Origin).Normalize()
    angle = atan2(vec.Y, vec.X)
    axis = Line.CreateBound(intersection_point, XYZ(intersection_point.X, intersection_point.Y, intersection_point.Z + 1))
    DB.ElementTransformUtils.RotateElement(doc, new_instance.Id, axis, angle)
    
    # Set family parameters including service name with error handling
    params = {
        'FP_Product Entry': 'Overall Size',
        'FP_Service Name': 'Fabrication Service Name',
        'FP_Service Abbreviation': 'Fabrication Service Abbreviation'
    }
    for fam_param, pipe_param in params.items():
        try:
            param_value = get_parameter_value_by_name_AsString(pipe, pipe_param)
            if param_value is not None:
                set_parameter_by_name(new_instance, fam_param, param_value)
            else:
                set_parameter_by_name(new_instance, fam_param, "")
                print("Warning: Pipe missing '{0}', set '{1}' to empty string".format(pipe_param, fam_param))
        except Exception as e:
            print("Error setting '{0}' from '{1}': {2}".format(fam_param, pipe_param, str(e)))
            set_parameter_by_name(new_instance, fam_param, "")
    
    return new_instance

def get_upper_level(current_view, all_levels):
    """Find the next highest level above the view's associated level by elevation"""
    view_level = current_view.GenLevel
    view_elevation = view_level.Elevation
    
    upper_level = None
    min_above_elevation = float('inf')
    
    for lvl in all_levels:
        if lvl.Elevation > view_elevation and lvl.Elevation < min_above_elevation:
            min_above_elevation = lvl.Elevation
            upper_level = lvl
    
    return upper_level

def main():
    """Main execution: place sleeves at vertical pipe-level intersections based on view type or selection"""
    family_symbol = load_family()
    if not family_symbol:
        print("Failed to load family symbol, load the family manually from here:\n{0}".format(family_path))
        return
    
    # Check for pre-selected elements
    selected_ids = uidoc.Selection.GetElementIds()
    all_levels = FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
    existing_sleeves = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
    
    with Transaction(doc, 'Place Sleeves at Intersections') as t:
        t.Start()
        if not family_symbol.IsActive:
            family_symbol.Activate()
            doc.Regenerate()
        
        placed_count = 0
        
        # Mode 1: Pre-selected pipe in Floor Plan - Place sleeve at upper level
        if selected_ids.Count > 0 and curview.ViewType == ViewType.FloorPlan:
            upper_level = get_upper_level(curview, all_levels)
            if not upper_level:
                print("No upper level found above the current floor plan view")
                t.Commit()
                return
            for element_id in selected_ids:
                pipe = doc.GetElement(element_id)
                if (pipe.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationPipework) and 
                    is_vertical_pipe(pipe)):
                    intersections = get_pipe_intersections(pipe, upper_level)
                    for point in intersections:
                        new_sleeve = place_sleeve_at_intersection(pipe, point, family_symbol, upper_level, existing_sleeves)
                        if new_sleeve:
                            placed_count += 1
        
        # Mode 2: Pre-selected pipe in 3D View - Place sleeves at visible level intersections
        elif selected_ids.Count > 0 and curview.ViewType == ViewType.ThreeD:
            # Get all visible fabrication pipes in the current 3D view
            visible_pipes = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework).WhereElementIsNotElementType().ToElements()
            visible_pipe_ids = {pipe.Id for pipe in visible_pipes}  # Set of visible pipe IDs
            
            # Get the section box of the 3D view (if active)
            section_box = curview.GetSectionBox() if curview.IsSectionBoxActive else None
            if section_box:
                min_point = section_box.Min
                max_point = section_box.Max
                # Transform coordinates to world space
                transform = section_box.Transform
                min_point = transform.OfPoint(min_point)
                max_point = transform.OfPoint(max_point)
            
            # Filter levels to those within the section box (if active)
            visible_levels = []
            for level in all_levels:
                if section_box:
                    level_z = level.Elevation
                    if min_point.Z <= level_z <= max_point.Z:  # Check if level is within section box Z-range
                        visible_levels.append(level)
                else:
                    visible_levels.append(level)  # If no section box, include all levels
            
            for element_id in selected_ids:
                if element_id not in visible_pipe_ids:  # Skip if the selected pipe isn't visible in the view
                    continue
                pipe = doc.GetElement(element_id)
                if (pipe.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationPipework) and 
                    is_vertical_pipe(pipe)):
                    for level in visible_levels:
                        intersections = get_pipe_intersections(pipe, level)
                        for point in intersections:
                            # Additional check: ensure intersection point is within section box
                            if section_box and not (min_point.X <= point.X <= max_point.X and 
                                                   min_point.Y <= point.Y <= max_point.Y and 
                                                   min_point.Z <= point.Z <= max_point.Z):
                                continue  # Skip if intersection is outside section box
                            new_sleeve = place_sleeve_at_intersection(pipe, point, family_symbol, level, existing_sleeves)
                            if new_sleeve:
                                placed_count += 1
        
        # Mode 3: 3D View without selection - Populate all visible intersections within section box
        elif curview.ViewType == ViewType.ThreeD:
            # Get all visible fabrication pipes in the current 3D view
            pipes = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework).WhereElementIsNotElementType().ToElements()
            
            # Get the section box of the 3D view (if active)
            section_box = curview.GetSectionBox() if curview.IsSectionBoxActive else None
            if section_box:
                min_point = section_box.Min
                max_point = section_box.Max
                # Transform coordinates to world space
                transform = section_box.Transform
                min_point = transform.OfPoint(min_point)
                max_point = transform.OfPoint(max_point)
            
            # Filter levels to those within the section box (if active)
            visible_levels = []
            for level in all_levels:
                if section_box:
                    level_z = level.Elevation
                    if min_point.Z <= level_z <= max_point.Z:  # Check if level is within section box Z-range
                        visible_levels.append(level)
                else:
                    visible_levels.append(level)  # If no section box, include all levels
            
            for pipe in pipes:
                if not is_vertical_pipe(pipe):
                    continue
                for level in visible_levels:
                    intersections = get_pipe_intersections(pipe, level)
                    for point in intersections:
                        # Additional check: ensure intersection point is within section box
                        if section_box and not (min_point.X <= point.X <= max_point.X and 
                                               min_point.Y <= point.Y <= max_point.Y and 
                                               min_point.Z <= point.Z <= max_point.Z):
                            continue  # Skip if intersection is outside section box
                        new_sleeve = place_sleeve_at_intersection(pipe, point, family_symbol, level, existing_sleeves)
                        if new_sleeve:
                            placed_count += 1
        
        # Mode 4: Floor Plan View without selection - Populate only upper level
        elif curview.ViewType == ViewType.FloorPlan:
            upper_level = get_upper_level(curview, all_levels)
            if not upper_level:
                print("No upper level found above the current view's level")
                t.Commit()
                return
            pipes = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework).WhereElementIsNotElementType().ToElements()
            for pipe in pipes:
                if not is_vertical_pipe(pipe):
                    continue
                intersections = get_pipe_intersections(pipe, upper_level)
                for point in intersections:
                    new_sleeve = place_sleeve_at_intersection(pipe, point, family_symbol, upper_level, existing_sleeves)
                    if new_sleeve:
                        placed_count += 1
        
        else:
            print("Script only runs in 3D or Floor Plan views, or with pre-selected pipes in supported views")
            t.Commit()
            return
        
        t.Commit()
        #print("Placed {0} sleeve instances at pipe-level intersections".format(placed_count))

if __name__ == '__main__':
    main()