
import Autodesk
from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource, Transaction, FilteredElementCollector, Family, TransactionGroup,\
                                BuiltInCategory, FamilySymbol, BuiltInParameter, Reference, IndependentTag, TagMode, TagOrientation

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'Fabrication Hanger - FP_Rod Size'
FamilyType = 'Fabrication Hanger - FP_Rod Size'
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

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)


family_pathCC1 = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES\Annotation\Fabrication Hanger - FP_Rod Size.rfa'

tg = TransactionGroup(doc, "Add Rod Size Tags")
tg.Start()

t = Transaction(doc, 'Load Rod Size Tag Family')
#Start Transaction
t.Start()
if Fam_is_in_project == False:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC1, fload_handler)
t.Commit()

Pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers)\
                                                                  .WhereElementIsNotElementType()
familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangerTags)\
                                                             .OfClass(FamilySymbol)\
                                                             .ToElements()

ItmList1 = list()

t = Transaction(doc, 'Tag Hanger Rods')
t.Start()
for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        if famtype.IsActive == False:
            famtype.Activate()
            doc.Regenerate()

for valve in Pipe_collector:
    ST = valve.ServiceType
    if ST == 56:
        R = Reference(valve)
        ValveLocation = valve.Origin
        IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, ValveLocation)
t.Commit()


# Creating collector instance and collecting all the fabrication hangers from the model
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

#End Transaction Group
tg.Assimilate()









