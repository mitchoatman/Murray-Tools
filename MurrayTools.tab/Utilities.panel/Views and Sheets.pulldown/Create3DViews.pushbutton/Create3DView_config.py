# coding: utf8
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import FilteredElementCollector, ViewFamilyType, ViewFamily, View3D, Transaction, Category

doc = revit.doc
output = script.get_output()
logger = script.get_logger()

# Select scope boxes (Volumes of Interest)
volumes_of_interest = forms.SelectFromList.show(
    FilteredElementCollector(doc).OfCategory(
        revit.DB.BuiltInCategory.OST_VolumeOfInterest
    ),
    "Select Scope Boxes for 3D views",
    multiselect=True,
    name_attr="Name",
)

# Get ViewFamilyType for 3D views
view_family_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).WhereElementIsElementType().ToElements()
view_family_type = next((vft for vft in view_family_types if vft.ViewFamily == ViewFamily.ThreeDimensional), None)

if view_family_type is None:
    forms.alert("No ViewFamilyType for 3D Views found.")
    script.exit()

def create_3d_views_with_scope_boxes(volumes_of_interest):
    # Collect existing 3D views to avoid duplicate names
    existing_views = FilteredElementCollector(doc).OfClass(View3D).ToElements()
    existing_view_names = {view.Name for view in existing_views}

    created_views = []
    skipped_views = []

    with Transaction(doc, "Create 3D Views with Scope Boxes") as t:
        t.Start()
        for voi in volumes_of_interest:
            view_name = voi.Name

            # Check if a view with the same name already exists
            if view_name in existing_view_names:
                skipped_views.append(view_name)
                logger.warning("View '{}' already exists and was skipped.".format(view_name))
                continue

            try:
                # Create new 3D isometric view
                new_view = View3D.CreateIsometric(doc, view_family_type.Id)
                new_view.Name = view_name

                # Assign the scope box to the view
                parameter = new_view.get_Parameter(
                    revit.DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP
                )
                if parameter:
                    parameter.Set(voi.Id)
                else:
                    logger.warning("Volume of interest parameter not found for view {}.".format(view_name))

                # Turn off scope boxes, grids, and levels categories
                categories_to_hide = [
                    revit.DB.BuiltInCategory.OST_VolumeOfInterest,
                    revit.DB.BuiltInCategory.OST_Grids,
                    revit.DB.BuiltInCategory.OST_Levels
                ]

                for category_id in categories_to_hide:
                    category = Category.GetCategory(doc, category_id)
                    if category:
                        # Check if view template controls visibility
                        view_template = new_view.ViewTemplateId
                        if view_template != revit.DB.ElementId.InvalidElementId:
                            # If view template is applied, modify the view's visibility directly
                            new_view.SetCategoryHidden(category.Id, True)
                        else:
                            # If no view template, modify visibility through the view's visibility settings
                            new_view.SetCategoryHidden(category.Id, True)

                created_views.append(view_name)
                logger.info("New 3D view created: {}".format(view_name))

            except Exception as e:
                logger.error("Failed to create view {}: {}".format(view_name, str(e)))

        t.Commit()

    # Print results
    # created_message = "\n".join("- {}".format(view) for view in created_views) if created_views else "None"
    # skipped_message = "\n".join("- {}".format(view) for view in skipped_views) if skipped_views else "None"
    # output.print_md("**Views Created:**\n{}\n\n**Views Skipped (Already Existing):**\n{}".format(created_message, skipped_message))

# Execute if scope boxes are selected
if volumes_of_interest:
    create_3d_views_with_scope_boxes(volumes_of_interest)
else:
    forms.alert("No scope boxes selected.")