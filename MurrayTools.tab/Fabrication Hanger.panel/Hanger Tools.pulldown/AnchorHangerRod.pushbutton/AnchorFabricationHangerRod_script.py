#Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FamilySymbol, Structure, Transaction, BuiltInParameter, \
                                Family, TransactionGroup, FamilyInstance, FabricationRodInfo
import os

DB = Autodesk.Revit.DB
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
CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")            
Hanger = [doc.GetElement(elId) for elId in pipesel]

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ExtentHangerRod.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

with open(filepath, 'r') as f:
    PrevInput = f.read()

def get_existing_rod_controls():
    # Collect all existing FP_Rod Control family instances in the document.
    rod_controls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation)\
                                               .WhereElementIsNotElementType()\
                                               .ToElements()
    existing_locations = {}
    for rc in rod_controls:
        if rc.Symbol.Family.Name == FamilyName and rc.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType:
            loc = rc.Location
            if isinstance(loc, DB.LocationPoint):
                existing_locations[rc.Id] = loc.Point
    return existing_locations

def is_location_occupied(target_point, existing_locations, tolerance=0.01):
    # Check if a target point is within tolerance of any existing Rod Control location.
    for loc in existing_locations.values():
        distance = target_point.DistanceTo(loc)
        if distance <= tolerance:  # Tolerance in feet
            return True
    return False

if len(Hanger) > 0:
    path, filename = os.path.split(__file__)
    NewFilename = '\FP_Rod Control.rfa'

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

    ItmList1 = []
    ItmList2 = []

    family_pathCC = path + NewFilename

    tg = TransactionGroup(doc, "Add Rod Control")
    tg.Start()

    t = Transaction(doc, 'Load Rod Control')
    t.Start()
    if not Fam_is_in_project:
        fload_handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(family_pathCC, fload_handler)
    t.Commit()

    # Get existing Rod Control locations
    existing_rod_locations = get_existing_rod_controls()

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
                rod_info = e.GetRodInfo()
                rod_count = rod_info.RodCount
                ItmList1.append(rod_count)
                for n in range(rod_count):
                    rodloc = rod_info.GetRodEndPosition(n)
                    ItmList2.append(rodloc)
            for hangerlocation in ItmList2:
                if not is_location_occupied(hangerlocation, existing_rod_locations):
                    familyInst = doc.Create.NewFamilyInstance(hangerlocation, famtype, Structure.StructuralType.NonStructural)
    t.Commit()

    t = Transaction(doc, 'Attach Rod to Structure')
    t.Start()
    for e in Hanger:
        rod_info = e.GetRodInfo()
        rod_count = rod_info.RodCount
        for n in range(rod_count):
            rod_info.AttachToStructure()
    t.Commit()

    tg.Assimilate()