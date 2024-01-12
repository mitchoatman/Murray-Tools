
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart, FabricationAncillaryUsage, Transaction, FilteredElementCollector, BuiltInCategory, TransactionGroup
from Autodesk.Revit.UI.Selection import *
from rpw.ui.forms import SelectFromList
from SharedParam.Add_Parameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return true

pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")            
hangers = [doc.GetElement( elId ) for elId in pipesel]

if len(hangers) > 0:

    value = SelectFromList('Select Rod Size', ['A - 3/8','B - 1/2','C - 5/8','D - 3/4','E - 7/8','F - 1','G - 1-1/4'])
    if value == 'A - 3/8':
        newrodkit = 58
    if value == 'B - 1/2':
        newrodkit = 42
    if value == 'C - 5/8':
        newrodkit = 31
    if value == 'D - 3/4':
        newrodkit = 62
    if value == 'E - 7/8':
        newrodkit = 64
    if value == 'F - 1':
        newrodkit = 67
    if value == 'G - 1-1/4':
        newrodkit = 70

    tg = TransactionGroup(doc, "Change Hanger Rod")
    tg.Start()

    t = Transaction(doc, "Set Hanger Rod")
    t.Start()
    for hanger in hangers:
        hanger.HangerRodKit = newrodkit
    t.Commit()

    t = Transaction(doc, "Update FP Parameter")
    t.Start()
    for x in hangers:
        [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
    t.Commit()
    
    #End Transaction Group
    tg.Assimilate()
else:
    print 'At least one fabrication hanger must be selected.'
