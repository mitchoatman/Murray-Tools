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
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString, set_parameter_by_name

path, filename = os.path.split(__file__)
NewFilename = r'\Duct Hanger Clearance.rfa'

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
FamilyName = 'Duct Hanger Clearance'
FamilyType = 'Duct Hanger Clearance'
Fam_is_in_project = any(f.Name == FamilyName for f in families)

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
    for loc in existing_locations.values():
        if target_point.DistanceTo(loc) <= tolerance:
            return True
    return False

def stretch_brace(familyInst, BraceOffsetZ):
    set_parameter_by_name(familyInst, "Length", .1666666)
    set_parameter_by_name(familyInst, "Width", .1666666)
    set_parameter_by_name(familyInst, "Height", BraceOffsetZ)

# Load the family if not present
tg = TransactionGroup(doc, "Hanger Clearance")
tg.Start()

t = Transaction(doc, 'Load Box Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()

# Get the FamilySymbol once
familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel)\
                                           .OfClass(FamilySymbol)\
                                           .ToElements()

# Place instances per hanger
t = Transaction(doc, 'Populate Rod Clearance')
t.Start()
famtype = None
for ft in familyTypes:
    typeName = ft.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if ft.Family.Name == FamilyName and typeName == FamilyType:
        famtype = ft
        famtype.Activate()
        doc.Regenerate()
        break

if famtype is None:
    raise Exception("Family type 'Duct Hanger Clearance' not found.")

# Get existing crop circle locations
existing_crop_locations = get_existing_crop_circles()

Hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers)\
                                                            .WhereElementIsNotElementType()
for e in Hanger_collector:
    ItmDims = e.GetDimensions()
    BraceOffsetZ = None
    for dta in ItmDims:
        if dta.Name in ['Rod Length', 'RodLength', 'Rod Extn Above', 'Top Length']:
            BraceOffsetZ = e.GetDimensionValue(dta)
            break  # Use the first valid rod length found

    if BraceOffsetZ is None:
        continue  # Skip if no rod length is found

    rod_info = e.GetRodInfo()
    rod_count = rod_info.RodCount
    for n in range(rod_count):
        rodloc = rod_info.GetRodEndPosition(n)
        if not is_location_occupied(rodloc, existing_crop_locations):
            familyInst = doc.Create.NewFamilyInstance(rodloc, famtype, Structure.StructuralType.NonStructural)
            stretch_brace(familyInst, BraceOffsetZ)
            # Add the new instance's location to prevent duplicates in this run
            existing_crop_locations[familyInst.Id] = rodloc

t.Commit()
tg.Assimilate()