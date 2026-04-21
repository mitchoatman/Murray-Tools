import Autodesk
from Autodesk.Revit.DB import Transaction

doc = __revit__.ActiveUIDocument.Document
DB = Autodesk.Revit.DB
curview = doc.ActiveView

t = Transaction(doc, "Reset Temporary Hide/Isolate")
t.Start()

curview.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)

t.Commit()