
#Imports
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol
from rpw.ui.forms import TextInput
from Autodesk.Revit.UI.Selection import *
import os

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_SetBearerExtn.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    f = open((filepath), 'w')
    f.write('1')
    f.close()

class MySelectionFilter(ISelectionFilter):
    def __init__(self):
        pass
    def AllowElement(self, element):
        if element.Category.Name == 'MEP Fabrication Hangers':
            return True
        else:
            return False
    def AllowReference(self, element):
        return False
selection_filter = MySelectionFilter()
Hanger = uidoc.Selection.PickElementsByRectangle(selection_filter)

if len(Hanger) > 0:

    f = open((filepath), 'r')
    PrevInput = f.read()
    f.close()

    #This displays dialog
    value = TextInput('New Bearer Extension (In Inches)', default = PrevInput)
    valuenum = (float(value) / 12)

    f = open((filepath), 'w')
    f.write(value)
    f.close()

    for x in range(2):

        t = Transaction(doc, 'Modify Bearer Extension')
        t.Start()

        for e in Hanger:
            STName = e.GetRodInfo().RodCount
            STName1 = e.GetRodInfo()
            for n in range(STName):
                STName1.SetBearerExtension(n, valuenum)

        t.Commit()
else:
    forms.alert('At least one fabrication hanger must be selected.')

