import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, FabricationPart, FabricationConfiguration
from SharedParam.Add_Parameters import Shared_Params
from pyrevit import DB, revit

Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

# FUNCTION TO SET PARAMETER VALUE
def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

# FUNCTION TO GET PARAMETER VALUE (change "AsDouble()" to "AsString()" to change data type.)
def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsString()

# Create a FilteredElementCollector to get all FabricationPart elements
AllElements = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

# Retrieve parameter values in advance
prefix_values = {}  # Dictionary to store prefix values for elements

for part in AllElements:
    ItemNumPrefix = part.LookupParameter('VM NUMBER PREFIX')
    if ItemNumPrefix and ItemNumPrefix.HasValue:
        prefix_values[part.Id] = ItemNumPrefix.AsString()

t = Transaction(doc, "Combine VM Numbering")
t.Start()

for part in AllElements:
    ItemNum = get_parameter_value_by_name(part, 'Item Number')
    if ItemNum:
        ItemNumPrefix = prefix_values.get(part.Id)
        if ItemNumPrefix and not ItemNum.startswith(ItemNumPrefix):
            set_parameter_by_name(part, 'Item Number', ItemNumPrefix + ItemNum)

t.Commit()
