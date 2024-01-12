#Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox
from pyrevit import script
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FamilySymbol, Structure, Transaction, BuiltInParameter, \
                                Family, TransactionGroup, FamilyInstance
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

    # Display dialog
    components = [
        Label('TOS Elevation from 0 *Input in this format FT-IN*'),
        TextBox('Elevation', PrevInput),
        CheckBox('checkboxvalue', 'Add Rod Control', default=True),
        Button('Ok')
        ]
    form = FlexForm('Modify Hanger Rod', components)
    form.show()


    # Convert dialog input into variable
    value = (form.values['Elevation'])
    InputFT = float(value.split("-", 1)[0])
    InputIN = (float(value.split("-", 1)[1]) / 12)
    valuenum = InputFT + InputIN
    RodControl = (form.values['checkboxvalue'])

    f = open((filepath), 'w')
    f.write(value)
    f.close()

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
else:
    print 'At least one fabrication hanger must be selected.'

        #print("Coordinates: {}".format(stringrodpos))
        #print("Length: {}".format(zcoord))




if RodControl == True:

    path, filename = os.path.split(__file__)
    NewFilename = '\FP_Rod Control.rfa'

    # Search project for all Families
    families = FilteredElementCollector(doc).OfClass(Family)
    # Set desired family name and type name:
    FamilyName = 'FP_Rod Control'
    FamilyType = 'FP_Rod Control'
    # Check if the family is in the project
    Fam_is_in_project = any(f.Name == FamilyName for f in families)
    #print("Family '{}' is in project: {}".format(FamilyName, is_in_project))

    class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
        def OnFamilyFound(self, familyInUse, overwriteParameterValues):
            overwriteParameterValues.Value = False
            return True


        def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
            source.Value = DB.FamilySource.Family
            overwriteParameterValues.Value = False
            return True


    ItmList1 = list()
    ItmList2 = list()

    family_pathCC = path + NewFilename

    tg = TransactionGroup(doc, "Add Rod Control")
    tg.Start()

    t = Transaction(doc, 'Load Rod Control')
    #Start Transaction
    t.Start()
    if Fam_is_in_project == False:
        fload_handler = FamilyLoaderOptionsHandler()
        family = doc.LoadFamily(family_pathCC, fload_handler)
    t.Commit()

    familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation)\
                                                .OfClass(FamilySymbol)\
                                                .ToElements()

    t = Transaction(doc, 'Populate Rod Control')
    t.Start()
    for famtype in familyTypes:
        typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if famtype.Family.Name == FamilyName and typeName == FamilyType:
            famtype.Activate()
            doc.Regenerate()
            for e in Hanger:
                STName = e.GetRodInfo().RodCount
                ItmList1.append(STName)
                STName1 = e.GetRodInfo()
                for n in range(STName):
                    rodloc = STName1.GetRodEndPosition(n)
                    ItmList2.append(rodloc)
            for hangerlocation in ItmList2:
                familyInst = doc.Create.NewFamilyInstance(hangerlocation, famtype, Structure.StructuralType.NonStructural)

    t.Commit()
    #End Transaction Group
    tg.Assimilate()