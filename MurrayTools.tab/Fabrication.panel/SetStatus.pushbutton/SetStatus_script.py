import Autodesk
from pyrevit import revit, forms
from Autodesk.Revit.DB import FabricationConfiguration, ElementId, Transaction
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Collections.Generic import HashSet

# Get document and UIDocument
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Get fabrication configuration and status list
config = FabricationConfiguration.GetFabricationConfiguration(doc)
status_names = []

if config:
    status_ids = config.GetAllPartStatuses()
    if status_ids and status_ids.Count > 0:
        for status_id in status_ids:
            name = config.GetPartStatusDescription(status_id)
            status_names.append(name)
    else:
        print "No statuses found in the fabrication configuration."
        status_names = ["None"]  # Fallback option
else:
    print "No fabrication configuration found in this project."
    status_names = ["None"]  # Fallback option

# Get all unique STRATUS Status values from elements in the current view
view = doc.ActiveView
collector = Autodesk.Revit.DB.FilteredElementCollector(doc, view.Id)
elements = collector.WhereElementIsNotElementType().ToElements()
stratus_statuses = set()

for elem in elements:
    param = elem.LookupParameter("STRATUS Status")
    if param and param.StorageType == Autodesk.Revit.DB.StorageType.String and param.HasValue:
        stratus_statuses.add(param.AsString())

stratus_statuses = list(stratus_statuses)
if not stratus_statuses:
    stratus_statuses = ["None"]  # Fallback if no statuses found

# Show dialog to select statuses to exclude
excluded_statuses = forms.SelectFromList.show(
    stratus_statuses,
    title="Select STRATUS Statuses to Exclude",
    button_name="Confirm",
    multiselect=True
)

if not excluded_statuses:
    print "No statuses excluded. All elements can be selected."
    excluded_statuses = []

# Custom selection filter to exclude elements with selected statuses
class StatusSelectionFilter(ISelectionFilter):
    def __init__(self, excluded_statuses):
        self.excluded_statuses = excluded_statuses

    def AllowElement(self, element):
        param = element.LookupParameter("STRATUS Status")
        if param and param.StorageType == Autodesk.Revit.DB.StorageType.String and param.HasValue:
            return param.AsString() not in self.excluded_statuses
        return True  # Allow elements without STRATUS Status or if status not excluded

    def AllowReference(self, reference, point):
        return True

# Prompt user to select elements with the custom filter
try:
    selection_filter = StatusSelectionFilter(excluded_statuses)
    selected_elements = uidoc.Selection.PickObjects(
        ObjectType.Element,
        selection_filter,
        "Select elements to assign status (excluded statuses filtered)"
    )
except Exception as e:
    print "Selection canceled or failed: %s" % str(e)
    import sys
    sys.exit()

if not selected_elements:
    print "No elements selected. Exiting."
    import sys
    sys.exit()

# Show dropdown dialog for status selection
selected_status = forms.SelectFromList.show(
    status_names,
    title="Select Fabrication Status",
    button_name="Confirm",
    multiselect=False
)
if not selected_status:
    print "No status selected. Exiting."
    import sys
    sys.exit()

# Apply the selected status to the selected elements
t = Transaction(doc, "Set STRATUS Status")
t.Start()
try:
    for ref in selected_elements:
        elem = doc.GetElement(ref.ElementId)
        param = elem.LookupParameter("STRATUS Status")
        if param and param.StorageType == Autodesk.Revit.DB.StorageType.String:
            param.Set(selected_status)
    t.Commit()
except Exception as e:
    t.RollBack()
    print "Failed to set STRATUS Status: %s" % str(e)