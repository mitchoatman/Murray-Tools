# coding: utf8
from pyrevit import revit, forms, script

doc = revit.doc
output = script.get_output()
logger = script.get_logger()

# Select scope boxes (Volumes of Interest)
volumes_of_interest = forms.SelectFromList.show(
    revit.DB.FilteredElementCollector(doc).OfCategory(
        revit.DB.BuiltInCategory.OST_VolumeOfInterest
    ),
    "Select Scope Boxes for dependent views",
    multiselect=True,
    name_attr="Name",
)

# Select views to duplicate as dependent views
views_to_duplicate = forms.select_views(use_selection=True)
if not views_to_duplicate:
    views_to_duplicate = forms.SelectFromList.show(
        revit.DB.FilteredElementCollector(doc).OfClass(revit.DB.View).ToElements(),
        "Select Views to Create Dependent Views",
        multiselect=True,
        name_attr="Name",
    )

def create_dependent_views(views_to_duplicate, volumes_of_interest):
    for view in views_to_duplicate:
        try:
            for voi in volumes_of_interest:
                # Duplicate the view as a dependent view
                new_view_id = view.Duplicate(revit.DB.ViewDuplicateOption.AsDependent)
                new_view = doc.GetElement(new_view_id)
                
                # Set the volume of interest to the new dependent view
                parameter = new_view.get_Parameter(
                    revit.DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP
                )
                if parameter:
                    parameter.Set(voi.Id)
                else:
                    logger.warning("Volume of interest parameter not found for view {view.Name}.")
                
                # Set the dependent view name to the combination of the original view name and scope box name
                new_view.Name = "{} - {}".format(view.Name, voi.Name)
                logger.info("New dependent view created: {}".format(new_view.Name))
        except AttributeError as e:
            logger.debug("Failed to duplicate view {}: {}".format(view.Name, str(e)))

# Create dependent views if selections are made
if volumes_of_interest and views_to_duplicate:
    with revit.Transaction("BatchCreateDependentViews"):
        create_dependent_views(views_to_duplicate, volumes_of_interest)
