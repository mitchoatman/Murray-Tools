import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from pyrevit import revit, DB, UI

#define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

selection = revit.get_selection()

# Creating collector instance and collecting all the fabrication hangers from the model
Hanger_collector = FilteredElementCollector(doc, curview.Id)\
    .OfCategory(BuiltInCategory.OST_FabricationHangers)\
    .WhereElementIsNotElementType()\
    .ToElements()

elementlist = [hanger.Id for hanger in Hanger_collector]

selection.set_to(elementlist)

