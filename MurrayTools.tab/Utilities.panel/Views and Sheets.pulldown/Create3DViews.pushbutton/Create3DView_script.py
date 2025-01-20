import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, Level, ViewFamilyType, ViewFamily, View3D, BoundingBoxXYZ, XYZ, Transaction
#from Autodesk.Revit.UI import TaskDialog


doc = __revit__.ActiveUIDocument.Document

def create_3d_view_per_level():
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()

    # Sort levels by elevation in ascending order
    levels = sorted(levels, key=lambda l: l.Elevation)

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
            bbox.Min = XYZ(-100, -100, level.Elevation)  # Level's exact elevation for Min Z
            
            # Determine the max elevation
            if i < len(levels) - 1:  # If there's a level above
                max_elevation = levels[i + 1].Elevation
            else:  # If no level above, add a default height (e.g., 10 units)
                max_elevation = level.Elevation + 10

            bbox.Max = XYZ(100, 100, max_elevation)
            view.SetSectionBox(bbox)

            created_views.append(view_name)  # Add to created list

        t.Commit()

    # Prepare the message for the dialog
    created_message = "\n".join("- %s" % view for view in created_views) if created_views else "None"
    skipped_message = "\n".join("- %s" % view for view in skipped_views) if skipped_views else "None"


    # Create a dialog box to show the results
    dialog_message = (
        "Views Created:\n{0}\n\n"
        "Views Skipped (Already Existing):\n{1}".format(created_message, skipped_message)
    )

    # Show the dialog
    #TaskDialog.Show("View Creation Status", dialog_message)
    print dialog_message

create_3d_view_per_level()
