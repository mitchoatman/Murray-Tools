
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FamilySymbol, Structure, Transaction, BuiltInParameter, \
                                FabricationConfiguration, Family, TransactionGroup
import os

path, filename = os.path.split(__file__)
NewFilename = '\Crop-Circle.rfa'

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'Crop-Circle'
FamilyType = 'Hanger'
# Check if the family is in the project
Fam_is_in_project = any(f.Name == FamilyName for f in families)
#print("Family '{}' is in project: {}".format(FamilyName, is_in_project))

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

selection = [doc.GetElement(id) for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]

if selection:
    tg = TransactionGroup(doc, "Add Crop Circles")
    tg.Start()

    t = Transaction(doc, 'Load Crop-Circle Family')
    #Start Transaction
    t.Start()
    if Fam_is_in_project == False:
        fload_handler = FamilyLoaderOptionsHandler()
        family = doc.LoadFamily(family_pathCC, fload_handler)
    t.Commit()

    familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel)\
                                                .OfClass(FamilySymbol)\
                                                .ToElements()

    t = Transaction(doc, 'Populate Crop Circles')
    t.Start()

    # Collect rod locations (ItmList2) from selection first
    ItmList2 = []
    for e in selection:
        isfabpart = e.LookupParameter("Fabrication Service")
        if isfabpart and e.ItemCustomId == 838:
            STName = e.GetRodInfo().RodCount
            STName1 = e.GetRodInfo()
            for n in range(STName):
                rodloc = STName1.GetRodEndPosition(n)
                ItmList2.append(rodloc)

    #print "Collected Rod Locations: {}".format(ItmList2)

    # Iterate through family types and place instances
    for famtype in familyTypes:
        typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if famtype.Family.Name == FamilyName and typeName == FamilyType:
            famtype.Activate()
            doc.Regenerate()
            for hangerlocation in ItmList2:
                familyInst = doc.Create.NewFamilyInstance(
                    hangerlocation, famtype, Structure.StructuralType.NonStructural
                )
                #print "Placed instance at: {}".format(hangerlocation)

    t.Commit()
    #End Transaction Group
    tg.Assimilate()

else:
    Hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers)\
                                                                .WhereElementIsNotElementType()

    tg = TransactionGroup(doc, "Add Crop Circles")
    tg.Start()

    t = Transaction(doc, 'Load Crop-Circle Family')
    #Start Transaction
    t.Start()
    if Fam_is_in_project == False:
        fload_handler = FamilyLoaderOptionsHandler()
        family = doc.LoadFamily(family_pathCC, fload_handler)
    t.Commit()

    familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel)\
                                                .OfClass(FamilySymbol)\
                                                .ToElements()

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
                familyInst = doc.Create.NewFamilyInstance(hangerlocation, famtype, Structure.StructuralType.NonStructural)

    t.Commit()
    #End Transaction Group
    tg.Assimilate()









