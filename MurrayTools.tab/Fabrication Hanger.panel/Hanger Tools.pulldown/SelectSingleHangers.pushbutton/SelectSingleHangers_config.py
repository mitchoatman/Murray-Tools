import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory
from pyrevit import revit
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
Shared_Params()

#define the active Revit application and document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

selection = revit.get_selection()

# Creating collector instance and collecting all the fabrication hangers from the model
hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

elementlist = [hanger.Id for hanger in hanger_collector if (hanger.GetRodInfo().RodCount) > 1]
selection.set_to(elementlist)