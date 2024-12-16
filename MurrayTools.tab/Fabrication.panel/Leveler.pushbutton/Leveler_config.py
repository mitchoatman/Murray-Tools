from Autodesk.Revit.DB import BoundingBoxXYZ, FilteredElementCollector, Transaction, BuiltInCategory, FabricationPart, Level, LogicalOrFilter, ElementCategoryFilter
import System
doc = __revit__.ActiveUIDocument.Document
curview = doc.ActiveView

# Collect all Level elements, storing their names, elevations, and IDs
level_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
level_elevations = {}
level_ids = {}
for level in level_collector:
    level_elevations[level.Name] = level.Elevation
    level_ids[level.Name] = level.Id

# Sort levels by elevation (ascending order)
sorted_levels = sorted(level_elevations.items(), key=lambda x: x[1])

# Function to get the center Z elevation of a fabrication part's bounding box
def get_center_point_z(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    center = (bBox.Max + bBox.Min) / 2
    return center.Z
# Function to get the bottom Z elevation of a fabrication part's bounding box
def get_bottom_point_z(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    return bBox.Min.Z  # Return the Z coordinate of the bottom point

# Category filters for Fabrication Pipework and Fabrication Ductwork
pipework_filter = ElementCategoryFilter(BuiltInCategory.OST_FabricationPipework)
ductwork_filter = ElementCategoryFilter(BuiltInCategory.OST_FabricationDuctwork)
combined_filter = LogicalOrFilter(pipework_filter, ductwork_filter)

# Collect all MEP Fabrication Pipework and Ductwork elements in the current view
fabrication_elements = FilteredElementCollector(doc, curview.Id) \
                       .OfClass(FabricationPart) \
                       .WherePasses(combined_filter) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

# Start transaction
t = Transaction(doc, "Assign Levels to Fabrication Parts")
t.Start()

# Iterate through fabrication parts to assign the closest level below
for elem in fabrication_elements:
    try:
        center_z = get_bottom_point_z(elem.Id)
        
        # Find the closest level below the center_z
        assigned_level_id = None
        for i, (level_name, elev) in enumerate(sorted_levels):
            if elev > center_z:
                # If no level was below center_z, fall back to the closest level above
                assigned_level_id = level_ids[sorted_levels[i-1][0]] if i > 0 else level_ids[level_name]
                break
            assigned_level_id = level_ids[level_name]  # Update with level directly below

        # Assign the level directly below or the closest above if none are below
        if assigned_level_id:
            elem.LookupParameter("Reference Level").Set(assigned_level_id)
            #print("Assigned level ID {assigned_level_id} to element {elem.Id}")

    except:
        pass

# Commit transaction
t.Commit()
