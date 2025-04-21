import Autodesk
from pyrevit import revit, forms
from Autodesk.Revit.DB import FabricationConfiguration, ElementId, Transaction
from Autodesk.Revit.DB.Fabrication import FabricationSaveJobOptions
from Autodesk.Revit.DB.FabricationPart import SaveAsFabricationJob
from Autodesk.Revit.UI.Selection import ObjectType
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

# Prompt user to select elements
selected_elements = uidoc.Selection.PickObjects(ObjectType.Element, "Select elements to export and assign status")
if not selected_elements:
    print "No elements selected. Exiting."
    import sys
    sys.exit()

element_ids = [elem.ElementId for elem in selected_elements]
id_set = HashSet[ElementId]()
for id in element_ids:
    id_set.Add(id)

# Show dropdown dialog for status selection
selected_status = forms.SelectFromList.show(status_names, 
                                           title="Select Fabrication Status", 
                                           button_name="Confirm", 
                                           multiselect=False)
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

if result:
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
                    else:
                        print "Element ID %s has no valid 'STRATUS Status' parameter." % elem.Id
                t.Commit()
                # print "STRATUS Status set to: %s" % selected_status
            except Exception, e:
                t.RollBack()
                print "Failed to set STRATUS Status: %s" % str(e)
    except Exception, e:
        print "Export failed: %s" % str(e)
else:
    print "Fabrication job saving canceled."