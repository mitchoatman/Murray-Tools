
import Autodesk
import clr
import System
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB.Fabrication import FabricationSaveJobOptions
from Autodesk.Revit.DB.FabricationPart import SaveAsFabricationJob
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
clr.AddReference("System.Core")
from System.Collections.Generic import HashSet

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import SaveFileDialog, DialogResult

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

selected_elements = uidoc.Selection.PickObjects(ObjectType.Element)
element_ids = [elem.ElementId for elem in selected_elements]

# Create a new HashSet[ElementId] object
id_set = HashSet[ElementId]()
# Add the selected IDs to the HashSet
for id in element_ids:
    id_set.Add(id)

options = FabricationSaveJobOptions()

# Prompt user for file name and location
save_dialog = SaveFileDialog()
save_dialog.Filter = "Fabrication Job Files (*.maj)|*.maj"
result = save_dialog.ShowDialog()

if result == DialogResult.OK:
    file_path = save_dialog.FileName
    SaveAsFabricationJob(doc, id_set, file_path, options)
else:
    print("Fabrication job saving canceled.")
