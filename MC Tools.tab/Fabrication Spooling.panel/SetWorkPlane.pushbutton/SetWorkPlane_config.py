import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog

# Get the current Revit document and UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Start a transaction to modify the view
t = Transaction(doc, "Disable Workplane Display")
t.Start()

try:
    # Hide the active workplane in the active view
    curview.HideActiveWorkPlane()
    t.Commit()
except Exception as e:
    t.RollBack()
    TaskDialog.Show("Error", "Failed to disable workplane display: {}".format(str(e)))