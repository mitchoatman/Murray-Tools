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
from Autodesk.Revit.UI import Selection, TaskDialog
import os
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString, set_parameter_by_name

path, filename = os.path.split(__file__)
NewFilename = r'\Duct Hanger Clearance.rfa'

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
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
    generic_elements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel)\
                                                .WhereElementIsNotElementType()\
                                                .ToElements()
    existing_locations = {}
    for cc in generic_elements:
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

# Transaction group for rollback on error
tg = TransactionGroup(doc, "Hanger Clearance")
try:
    tg.Start()

    # Load family with error handling
    t = Transaction(doc, 'Load Box Family')
    try:
        t.Start()
        if not Fam_is_in_project:
            if os.path.exists(family_pathCC):
                fload_handler = FamilyLoaderOptionsHandler()
                success = doc.LoadFamily(family_pathCC, fload_handler)
                if not success:
                    raise Exception("Failed to load family file")
            else:
                raise Exception("Family file not found at " + family_pathCC)
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Family load failed: " + str(e))
        raise

    # Get the FamilySymbol within a transaction
    t = Transaction(doc, 'Get Family Symbol')
    try:
        t.Start()
        familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel)\
                                                   .OfClass(FamilySymbol)\
                                                   .ToElements()
        famtype = None
        for ft in familyTypes:
            typeName = ft.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if ft.Family.Name == FamilyName and typeName == FamilyType:
                famtype = ft
                famtype.Activate()
                break
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Family symbol retrieval failed: " + str(e))
        raise

    if famtype is None:
        TaskDialog.Show("Error", "Family type 'Duct Hanger Clearance' not found")
        raise Exception("Family type 'Duct Hanger Clearance' not found")

    # Place instances per hanger with batch processing
    t = Transaction(doc, 'Populate Rod Clearance')
    try:
        t.Start()
        existing_crop_locations = get_existing_crop_circles()
        Hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers)\
                                                                    .WhereElementIsNotElementType()\
                                                                    .ToElements()
        batch_size = 5
        hangers = list(Hanger_collector)
        for i in range(0, len(hangers), batch_size):
            batch = hangers[i:i + batch_size]
            for e in batch:
                ItmDims = e.GetDimensions()
                BraceOffsetZ = None
                for dta in ItmDims:
                    if dta.Name in ['Rod Length', 'RodLength', 'Rod Extn Above', 'Top Length']:
                        BraceOffsetZ = e.GetDimensionValue(dta)
                        break

                if BraceOffsetZ is None:
                    continue

                rod_info = e.GetRodInfo()
                rod_count = rod_info.RodCount
                for n in range(rod_count):
                    rodloc = rod_info.GetRodEndPosition(n)
                    if not is_location_occupied(rodloc, existing_crop_locations):
                        familyInst = doc.Create.NewFamilyInstance(rodloc, famtype, Structure.StructuralType.NonStructural)
                        stretch_brace(familyInst, BraceOffsetZ)
                        existing_crop_locations[familyInst.Id] = rodloc
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Instance placement failed: " + str(e))
        raise

    tg.Assimilate()
except Exception as e:
    tg.RollBack()
    TaskDialog.Show("Error", "Operation failed: " + str(e))