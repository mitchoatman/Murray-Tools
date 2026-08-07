import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (FilteredElementCollector, Level, ViewFamilyType, ViewFamily, View3D, BoundingBoxXYZ,
    XYZ, Transaction, Transform, ProjectLocation)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Get the bounding box of the active view (already in internal feet)
active_view_bbox = curview.CropBox

x_min = active_view_bbox.Min.X
x_max = active_view_bbox.Max.X
y_min = active_view_bbox.Min.Y
y_max = active_view_bbox.Max.Y

def get_level_elevation(level):
    """
    Get the true internal project elevation of the level, 
    matching the logic used in your sleeve placement script.
    """
    # ProjectElevation gives the internal coordinate Z height reliably
    if hasattr(level, 'ProjectElevation'):
        return level.ProjectElevation
    return level.Elevation

def create_3d_view_per_level():
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()

    # Sort levels by project elevation in ascending order
    levels = sorted(levels, key=lambda l: get_level_elevation(l))

    view_family_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).WhereElementIsElementType().ToElements()
    view_family_type = next((vft for vft in view_family_types if vft.ViewFamily == ViewFamily.ThreeDimensional), None)

    if view_family_type is None:
        raise ValueError("No ViewFamilyType for 3D Views found.")

    # Collect existing 3D views
    existing_views = FilteredElementCollector(doc).OfClass(View3D).ToElements()
    existing_view_names = {view.Name for view in existing_views}

    created_views = []
    skipped_views = []

    with Transaction(doc, "Create 3D Views per Level") as t:
        t.Start()
        for i, level in enumerate(levels):
            view_name = "Level {} 3D View".format(level.Name)

            # Check if a view with the same name already exists
            if view_name in existing_view_names:
                skipped_views.append(view_name)
                continue

            # Create the 3D view
            view = View3D.CreateIsometric(doc, view_family_type.Id)
            view.Name = view_name

            # Set the section box using accurate internal project elevations
            bbox = BoundingBoxXYZ()
            z_min = get_level_elevation(level)
            
            if i < len(levels) - 1:  # If there's a level above
                z_max = get_level_elevation(levels[i + 1])
            else:  # If no level above, add a default height buffer of 10 feet
                z_max = z_min + 10.0  

            bbox.Min = XYZ(x_min, y_min, z_min)
            bbox.Max = XYZ(x_max, y_max, z_max)
            view.SetSectionBox(bbox)

            created_views.append(view_name)

        t.Commit()

    # Prepare messages
    created_message = "\n".join("- %s" % view for view in created_views) if created_views else "None"
    skipped_message = "\n".join("- %s" % view for view in skipped_views) if skipped_views else "None"

    print("Views Created:\n{0}\n".format(created_message))
    print("Views Skipped (Already Existing):\n{0}".format(skipped_message))

create_3d_view_per_level()