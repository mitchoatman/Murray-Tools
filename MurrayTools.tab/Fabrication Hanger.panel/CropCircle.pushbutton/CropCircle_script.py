import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, 
    BuiltInCategory, 
    FamilySymbol, 
    Structure, 
    Transaction, 
    BuiltInParameter,
    FabricationConfiguration, 
    Family, 
    TransactionGroup
)
import os

path, filename = os.path.split(__file__)
NewFilename = r'\Crop-Circle.rfa'

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'Crop-Circle'
FamilyType = 'Hanger'
Fam_is_in_project = any(f.Name == FamilyName for f in families)

ItmList1 = []
ItmList2 = []

family_pathCC = path + NewFilename

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

def get_existing_crop_circles():
    # Collect all existing Crop-Circle family instances in the document.
    crop_circles = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel)\
                                                .WhereElementIsNotElementType()\
                                                .ToElements()
    existing_locations = {}
    for cc in crop_circles:
        if cc.Symbol.Family.Name == FamilyName and cc.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType:
            loc = cc.Location
            if isinstance(loc, DB.LocationPoint):
                existing_locations[cc.Id] = loc.Point
    return existing_locations

def is_location_occupied(target_point, existing_locations, tolerance=0.01):
    # Check if a target point is within tolerance of any existing Crop-Circle location.
    for loc in existing_locations.values():
        distance = target_point.DistanceTo(loc)
        if distance <= tolerance:  # Tolerance in feet
            return True
    return False


Hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers)\
                                                            .WhereElementIsNotElementType()

tg = TransactionGroup(doc, "Add Crop Circles")
tg.Start()

t = Transaction(doc, 'Load Crop-Circle Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()

familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel)\
                                            .OfClass(FamilySymbol)\
                                            .ToElements()

# Get existing Crop-Circle locations
existing_crop_locations = get_existing_crop_circles()

t = Transaction(doc, 'Populate Crop Circles')
t.Start()
for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        famtype.Activate()
        doc.Regenerate()
        for e in Hanger_collector:
            STName = e.GetRodInfo().RodCount
            ItmList1.append(STName)
            STName1 = e.GetRodInfo()
            for n in range(STName):
                rodloc = STName1.GetRodEndPosition(n)
                ItmList2.append(rodloc)
        for hangerlocation in ItmList2:
            if not is_location_occupied(hangerlocation, existing_crop_locations):
                familyInst = doc.Create.NewFamilyInstance(
                    hangerlocation, famtype, Structure.StructuralType.NonStructural
                )

t.Commit()
tg.Assimilate()