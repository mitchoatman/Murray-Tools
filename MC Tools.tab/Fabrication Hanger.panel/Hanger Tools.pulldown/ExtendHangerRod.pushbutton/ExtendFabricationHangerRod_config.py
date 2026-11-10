# Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FamilySymbol, Structure, Transaction, BuiltInParameter, \
                                Family, TransactionGroup, FamilyInstance
import os
import sys
import json

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return True

pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers to Extend")
Hanger = [doc.GetElement(elId) for elId in pipesel]

# Set up JSON File Path
folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ExtendHangerRod.json')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# Default Settings
default_settings = {
    'elevation': '5-6',
    'add_rod_control': False
}

# Ensure file exists with valid JSON content
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        json.dump(default_settings, f, indent=4)

# Read settings from JSON
try:
    with open(filepath, 'r') as f:
        config_data = json.load(f)
        PrevInput = config_data.get('elevation', '5-6')
        PrevRodControl = config_data.get('add_rod_control', False)
except Exception:
    PrevInput = '5-6'
    PrevRodControl = False

if len(Hanger) > 0:
    # Display dialog using saved JSON parameters
    components = [
        Label('TOS Elevation from 0 *Input in this format FT-IN*'),
        TextBox('Elevation', PrevInput),
        CheckBox('checkboxvalue', 'Add Rod Control', default=PrevRodControl),
        Button('Ok')
    ]
    form = FlexForm('Modify Hanger Rod', components)

    if not form.show():
        sys.exit()

    # Convert dialog input into variables
    value = form.values['Elevation']
    InputFT = float(value.split("-", 1)[0])
    InputIN = float(value.split("-", 1)[1]) / 12
    valuenum = InputFT + InputIN
    RodControl = form.values['checkboxvalue']

    # Update JSON config file with new UI values
    updated_settings = {
        'elevation': value,
        'add_rod_control': RodControl
    }
    with open(filepath, 'w') as f:
        json.dump(updated_settings, f, indent=4)

    ItmList1 = list()
    zLocs = []

    t = Transaction(doc, 'Extend Hanger Rods')
    t.Start()

    for e in Hanger:
        STName = e.GetRodInfo().RodCount
        ItmList1.append(STName)
        # Detaches rods from structure
        hgrhost = e.GetRodInfo().CanRodsBeHosted = False
        STName1 = e.GetRodInfo()
        for n in range(STName):
            rodlen = STName1.GetRodLength(n)
            rodpos = STName1.GetRodEndPosition(n)
            # Directly access Z-coordinate from XYZ object
            zcoord = rodpos.Z
            STName1.SetRodLength(n, rodlen + (valuenum - zcoord))
    t.Commit()
else:
    print('At least one fabrication hanger must be selected.')

if RodControl:
    path, filename = os.path.split(__file__)
    NewFilename = '\\FP_Rod Control.rfa'

    # Search project for all Families
    families = FilteredElementCollector(doc).OfClass(Family)
    # Set desired family name and type name:
    FamilyName = 'FP_Rod Control'
    FamilyType = 'FP_Rod Control'
    # Check if the family is in the project
    Fam_is_in_project = any(f.Name == FamilyName for f in families)

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
    t.Start()
    if not Fam_is_in_project:
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

    tg.Assimilate()

    # Run Attach to Structure in a new transaction
    t = Transaction(doc, 'Attach Hangers to Structure')
    t.Start()
    for e in Hanger:
        e.GetRodInfo().CanRodsBeHosted = True  # Re-enable hosting
        e.GetRodInfo().AttachToStructure()  # Attach to structure
    t.Commit()