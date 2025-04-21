import Autodesk
from pyrevit import revit
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Windows.Forms import SaveFileDialog, DialogResult
import os

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

selected_elements = uidoc.Selection.PickObjects(ObjectType.Element)
element_ids = [elem.ElementId for elem in selected_elements]

# Define the path for saving the last export location
folder_name = "C:\\Temp"
filepath = os.path.join(folder_name, "Ribbon_Exports.txt")

# Ensure the C:\Temp folder exists
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# Get the default Desktop path using %USERPROFILE%
default_desktop_path = os.path.expandvars("%USERPROFILE%\\Desktop")

# Load the last saved path from the text file, default to Desktop if not found
if os.path.exists(filepath):
    with open(filepath, 'r') as f:
        last_save_path = f.read().strip()
    if not os.path.exists(last_save_path):
        last_save_path = default_desktop_path
else:
    last_save_path = default_desktop_path

# File save dialog for PCF export using .NET
save_dialog = SaveFileDialog()
save_dialog.Title = "Save PCF File"
save_dialog.Filter = "Fabrication Job Files (*.pcf)|*.pcf"
save_dialog.DefaultExt = "pcf"
save_dialog.InitialDirectory = last_save_path
save_dialog.FileName = doc.Title  # Set default filename to project filename

if save_dialog.ShowDialog() == DialogResult.OK:
    file_path = save_dialog.FileName
    folder_path = os.path.dirname(file_path)

    # Export to PCF
    try:
        Autodesk.Revit.DB.Fabrication.FabricationUtils.ExportToPCF(doc, element_ids, file_path)
        # Save the selected folder path to the text file after successful export
        with open(filepath, 'w') as f:
            f.write(folder_path)
    except Exception as e:
        print("Export failed: " + str(e))
else:
    print("PCF file save canceled.")