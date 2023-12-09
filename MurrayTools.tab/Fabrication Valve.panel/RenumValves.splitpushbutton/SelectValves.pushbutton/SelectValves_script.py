import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from pyrevit import revit


#define the active Revit application and document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

selection = revit.get_selection()

# Creating collector instance and collecting all the fabrication hangers from the model
pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                   .WhereElementIsNotElementType() \
                   .ToElements()


elementlist = [valve.Id for valve in pipe_collector if valve.ServiceType == 53]
selection.set_to(elementlist)


