import Autodesk
from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource, Transaction, FilteredElementCollector, Family, TransactionGroup, \
                              BuiltInCategory, FamilySymbol, BuiltInParameter, Reference, IndependentTag, TagMode, TagOrientation, \
                              ViewType, View3D
from Parameters.Add_SharedParameters import Shared_Params
from Autodesk.Revit.UI import TaskDialog
import os
import sys

Shared_Params()

path, filename = os.path.split(__file__)
NewFilename = '\Fabrication Hanger - FP_Rod Size.rfa'

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Handle 3D view locking with validation
if curview.ViewType == ViewType.ThreeD:
    v3d = curview  # Safe cast as ViewType confirms View3D
    if v3d.IsPerspective:
        TaskDialog.Show("Error", "Tagging is not supported in perspective 3D views.")
        sys.exit()
    if "}" in v3d.Name:
        TaskDialog.Show("Error", "Tagging only allowed in Locked 3D view.  Your view contains invalid characters, rename the view if you wish to tag in 3D")
        sys.exit()
    
    t_lock = Transaction(doc, "Lock 3D View Orientation")
    t_lock.Start()
    try:
        v3d.SaveOrientationAndLock()
        t_lock.Commit()
        if not v3d.IsLocked:
            raise Exception("Lock verification failed")
    except:
        t_lock.RollBack()
        TaskDialog.Show("Error", "Failed to lock 3D view orientation. Ensure the view is not a template and retry.")
        sys.exit()
elif curview.ViewType not in [ViewType.FloorPlan, ViewType.AreaPlan, ViewType.Section]:
    TaskDialog.Show("Error", "This script can only run in a Floor Plan or Section view.")
    sys.exit()

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'Fabrication Hanger - FP_Rod Size'
FamilyType = 'Fabrication Hanger - FP_Rod Size'
Fam_is_in_project = any(f.Name == FamilyName for f in families)

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

family_pathCC1 = path + NewFilename

tg = TransactionGroup(doc, "Add Rod Size Tags")
tg.Start()

t = Transaction(doc, 'Load Rod Size Tag Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_pathCC1, fload_handler)
t.Commit()

Pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                                                          .WhereElementIsNotElementType()
familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangerTags) \
                                           .OfClass(FamilySymbol) \
                                           .ToElements()

t = Transaction(doc, 'Tag Hanger Rods')
t.Start()
for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        if not famtype.IsActive:
            famtype.Activate()
            doc.Regenerate()

for valve in Pipe_collector:
    ST = valve.ServiceType
    if ST == 56:
        R = Reference(valve)
        ValveLocation = valve.Origin
        IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, ValveLocation)
t.Commit()

hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

t = Transaction(doc, 'Sync Hanger Data')
t.Start()
try:
    list(map(lambda hanger: [set_parameter_by_name(hanger, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in hanger.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0], hanger_collector))
except:
    pass
t.Commit()

tg.Assimilate()