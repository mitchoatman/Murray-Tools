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

t = Transaction(doc, "Select Single Hangers")
t.Start()

for hanger in hanger_collector:
    try:
        if (hanger.GetRodInfo().RodCount) > 1:
            ItmDims = hanger.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Width':
                    TrapWidth = hanger.GetDimensionValue(dta)
                if dta.Name == 'Bearer Extn':
                    TrapExtn = hanger.GetDimensionValue(dta)
                if dta.Name == 'Right Rod Offset':
                    TrapRRod = hanger.GetDimensionValue(dta)
                if dta.Name == 'Left Rod Offset':
                    TrapLRod = hanger.GetDimensionValue(dta)
            BearerLength = TrapWidth + TrapExtn + TrapExtn
            set_parameter_by_name(hanger, 'FP_Bearer Length', BearerLength)
    except:
        pass

t.Commit()

elementlist = [hanger.Id for hanger in hanger_collector if not hanger.LookupParameter('FP_Bearer Length').HasValue]
selection.set_to(elementlist)