import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (FilteredElementCollector, Level, ViewFamilyType, ViewFamily, View3D, BoundingBoxXYZ,
    XYZ, Transaction, UnitUtils, UnitTypeId, Transform)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Get the bounding box of the active view
active_view_bbox = curview.CropBox

# Extract X and Y values from the active view's bounding box
active_view_min = active_view_bbox.Min
active_view_max = active_view_bbox.Max

x_min = active_view_min.X
x_max = active_view_bbox.Max.X
y_min = active_view_bbox.Min.Y
y_max = active_view_bbox.Max.Y

def get_elevation_from_survey(level):
    """
    Get the elevation of the level relative to the survey point.
    """
    project_location = doc.ActiveProjectLocation
    total_transform = project_location.GetTotalTransform()  # Transformation from Project Base to Survey
    elevation_in_project = level.Elevation
    elevation_in_survey = elevation_in_project + total_transform.Origin.Z  # Add Z-offset from transformation
    return elevation_in_survey

def convert_to_feet(value):
    """Convert project units to feet if necessary."""
    return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Feet)

def create_3d_view_per_level():
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()

    # Sort levels by survey elevation in ascending order
    levels = sorted(levels, key=lambda l: get_elevation_from_survey(l))

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
                skipped_views.append(view_name)  # Add to skipped list
                continue  # Skip creation if the view already exists

            # Create the 3D view
            view = View3D.CreateIsometric(doc, view_family_type.Id)
            view.Name = view_name

            # Set the section box for the view
            bbox = BoundingBoxXYZ()
            z_min = convert_to_feet(get_elevation_from_survey(level))
            if i < len(levels) - 1:  # If there's a level above
                z_max = convert_to_feet(get_elevation_from_survey(levels[i + 1]))
            else:  # If no level above, add a default height (e.g., 10 units)
                z_max = z_min + 10  # Add a buffer of 10 feet

            bbox.Min = XYZ(x_min, y_min, z_min)
            bbox.Max = XYZ(x_max, y_max, z_max)
            view.SetSectionBox(bbox)

            # # Print debug information for troubleshooting
            # print "View '{0}': SectionBox Min({1}, {2}, {3}) Max({4}, {5}, {6})".format(
                # view_name, x_min, y_min, z_min, x_max, y_max, z_max
            # )

            created_views.append(view_name)  # Add to created list

        t.Commit()

    # Prepare the message for the dialog
    created_message = "\n".join("- %s" % view for view in created_views) if created_views else "None"
    skipped_message = "\n".join("- %s" % view for view in skipped_views) if skipped_views else "None"

    # Print the results
    print "Views Created:\n{0}\n".format(created_message)
    print "Views Skipped (Already Existing):\n{0}".format(skipped_message)

create_3d_view_per_level()
