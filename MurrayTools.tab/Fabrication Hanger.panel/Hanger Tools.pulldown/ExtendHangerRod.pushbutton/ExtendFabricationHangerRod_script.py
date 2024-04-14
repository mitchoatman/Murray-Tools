#Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from rpw.ui.forms import CommandLink, TaskDialog
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FamilySymbol, Structure, Transaction, BuiltInParameter, \
                                Family, TransactionGroup, FamilyInstance, ReferencePlane
import os

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

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
CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers to Extend")            
Hanger = [doc.GetElement( elId ) for elId in pipesel]

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ExtentHangerRod.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    f = open((filepath), 'w')
    f.write('1')
    f.close()

f = open((filepath), 'r')
PrevInput = f.read()
f.close()

if len(Hanger) > 0:

    # Prompt the user to select an existing reference plane

    try:
        ref_plane = uidoc.Selection.PickObject(ObjectType.Element, "Select a reference plane")
        ref_plane = doc.GetElement(ref_plane.ElementId)
        if not isinstance(ref_plane, DB.ReferencePlane):
            ref_plane = None

        # Retrieve the reference plane elevation
        valuenum = ref_plane.FreeEnd.Z

        ItmList1 = list()
        zLocs = []

        t = Transaction(doc, 'Extend Hanger Rods')
        t.Start()

        for e in Hanger:
            STName = e.GetRodInfo().RodCount
            ItmList1.append(STName)
            #Detaches rods from structure
            hgrhost = e.GetRodInfo().CanRodsBeHosted = False
            STName1 = e.GetRodInfo()
            for n in range(STName):
                rodlen = STName1.GetRodLength(n)
                rodpos = STName1.GetRodEndPosition(n)
                #Turns rodpos into string and removes ( ) to clean it up.
                stringrodpos = str(rodpos).replace ('(', '').replace(')', '')
                length = len(stringrodpos)

                #Looks for "," to locate where Z coordinate starts
                zcoordloc = stringrodpos.rfind(', ', 0, length)
                #Removes x and y coordinate data and returns only z converted back to number.
                zcoord = float((stringrodpos[zcoordloc+2:length]))
                STName1.SetRodLength(n, rodlen + (valuenum - zcoord))

        t.Commit()
    except:
        pass
else:
    print 'At least one fabrication hanger must be selected.'
