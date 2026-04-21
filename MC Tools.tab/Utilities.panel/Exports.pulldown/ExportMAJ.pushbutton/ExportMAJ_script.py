import Autodesk
from pyrevit import revit, forms
from Autodesk.Revit.DB import FabricationConfiguration, ElementId, Transaction
from Autodesk.Revit.DB.Fabrication import FabricationSaveJobOptions
from Autodesk.Revit.DB.FabricationPart import SaveAsFabricationJob
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Collections.Generic import HashSet
from System.Windows.Forms import SaveFileDialog, DialogResult
import os

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
        "Select elements to export and assign status (excluded statuses filtered)"
    )
except Exception as e:
    print "Selection canceled or failed: %s" % str(e)
    import sys
    sys.exit()

if not selected_elements:
    print "No elements selected. Exiting."
    import sys
    sys.exit()

element_ids = [elem.ElementId for elem in selected_elements]
id_set = HashSet[ElementId]()
for id in element_ids:
    id_set.Add(id)

# Show dropdown dialog for status selection
selected_status = forms.SelectFromList.show(
    status_names,
    title="Select Fabrication Status",
    button_name="Confirm",
    multiselect=False
)
if not selected_status:
    print "No status selected. Proceeding with export only."
    selected_status = None  # Handle cancellation

# MAJ export logic
options = FabricationSaveJobOptions()
folder_name = "C:\\Temp"
filepath = os.path.join(folder_name, "Ribbon_Exports.txt")

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

default_desktop_path = os.path.expandvars("%USERPROFILE%\\Desktop")
if os.path.exists(filepath):
    with open(filepath, 'r') as f:
        last_save_path = f.read().strip()
    if not os.path.exists(last_save_path):
        last_save_path = default_desktop_path
else:
    last_save_path = default_desktop_path

save_dialog = SaveFileDialog()
save_dialog.Filter = "Fabrication Job Files (*.maj)|*.maj"
save_dialog.DefaultExt = "maj"
save_dialog.InitialDirectory = last_save_path
save_dialog.FileName = doc.Title
save_dialog.Title = "Save MAJ File"

result = save_dialog.ShowDialog()

if result == DialogResult.OK:
    file_path = save_dialog.FileName
    folder_path = os.path.dirname(file_path)
    
    # Export to MAJ
    try:
        SaveAsFabricationJob(doc, id_set, file_path, options)
        # print "MAJ file exported successfully to: %s" % file_path
        
        # Save the folder path
        with open(filepath, 'w') as f:
            f.write(folder_path)
        
        # If a status was selected, update the parameter
        if selected_status:
            t = Transaction(doc, "Set STRATUS Status")
            t.Start()
            try:
                for ref in selected_elements:
                    elem = doc.GetElement(ref.ElementId)
                    param = elem.LookupParameter("STRATUS Status")
                    if param and param.StorageType == Autodesk.Revit.DB.StorageType.String:
                        param.Set(selected_status)
                    # else:
                        # print "Element ID %s has no valid 'STRATUS Status' parameter." % elem.Id
                t.Commit()
                # print "STRATUS Status set to: %s" % selected_status
            except Exception as e:
                t.RollBack()
                print "Failed to set STRATUS Status: %s" % str(e)
    except Exception as e:
        print "Export failed: %s" % str(e)
else:
    print "Fabrication job saving canceled."