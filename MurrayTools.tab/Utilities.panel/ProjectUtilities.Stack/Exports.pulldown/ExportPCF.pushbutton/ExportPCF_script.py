import Autodesk
from pyrevit import revit, forms
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Windows.Forms import SaveFileDialog, DialogResult

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

selected_elements = uidoc.Selection.PickObjects(ObjectType.Element)
element_ids = [elem.ElementId for elem in selected_elements]

# Prompt user for file name and location
save_dialog = SaveFileDialog()
save_dialog.Filter = "Fabrication Job Files (*.pcf)|*.pcf"
result = save_dialog.ShowDialog()

if result == DialogResult.OK:
    file_path = save_dialog.FileName
    Autodesk.Revit.DB.Fabrication.FabricationUtils.ExportToPCF(doc, element_ids, file_path)
else:
    print("PCF file save canceled.")
