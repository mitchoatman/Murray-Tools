import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FilterStringLessOrEqual, FilterStringRule, \
ParameterValueProvider, ElementId, FilterStringBeginsWith, Transaction, FilterStringEquals, \
ElementParameterFilter, ParameterValueProvider, LogicalOrFilter, TransactionGroup, FabricationPart
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)



#FUNCTION TO SET PARAMETER VALUE
def set_parameter_by_name(element, parameterName, value):
	element.LookupParameter(parameterName).Set(value)

Strut_Collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_StructuralFraming) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

t = Transaction(doc, "Transfer Elevation")
t.Start()

for strut in Strut_Collector:
    try:
        elevation = strut.LookupParameter('Offset from Host').AsValueString()
        set_parameter_by_name(strut, 'Comments', elevation)
    except:
        pass

t.Commit()