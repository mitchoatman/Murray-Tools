import Autodesk
from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource, Transaction, FilteredElementCollector, Family, TransactionGroup,\
                                BuiltInCategory, FamilySymbol, BuiltInParameter, Reference, IndependentTag, TagMode, TagOrientation
import os
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

path, filename = os.path.split(__file__)
NewFilename = '\Pipe Accessory - Trimble Sleeve Size Tag.rfa'

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

try:
    # Search project for all Families
    families = FilteredElementCollector(doc).OfClass(Family)
    FamilyName = 'Pipe Accessory - Trimble Sleeve Size Tag'
    FamilyType = 'Pipe Accessory Tag - SLV Tag'
    Fam_is_in_project = any(f.Name == FamilyName for f in families)

    class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
        def OnFamilyFound(self, familyInUse, overwriteParameterValues):
            overwriteParameterValues.Value = False
            return True

        def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
            source.Value = DB.FamilySource.Family
            overwriteParameterValues.Value = False
            return True

    family_pathCC = path + NewFilename

    tg = TransactionGroup(doc, "Add Sleeve Tags")
    try:
        tg.Start()

        t = Transaction(doc, 'Load Sleeve Size Family')
        try:
            t.Start()
            if not Fam_is_in_project:
                fload_handler = FamilyLoaderOptionsHandler()
                if not os.path.exists(family_pathCC):
                    raise Exception("Family file not found: {}".format(family_pathCC))
                family = doc.LoadFamily(family_pathCC, fload_handler)
            t.Commit()
        except Exception as e:
            t.RollBack()
            raise Exception("Failed to load family: {}".format(str(e)))

        try:
            familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessoryTags)\
                                                    .OfClass(FamilySymbol)\
                                                    .ToElements()
        except Exception as e:
            raise Exception("Failed to collect family types: {}".format(str(e)))

        t = Transaction(doc, 'Tag Sleeves')
        try:
            t.Start()
            for famtype in familyTypes:
                typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                if famtype.Family.Name == FamilyName and typeName == FamilyType:
                    if not famtype.IsActive:
                        famtype.Activate()
                        doc.Regenerate()

            accessory_models_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PipeAccessory)
            accessory_elements = [element for element in accessory_models_collector 
                               if "Metal Sleeve" in element.Name or 
                                  "Plastic Sleeve" in element.Name or 
                                  "Round Floor Sleeve" in element.Name or 
                                  "Cast Iron Sleeve" in element.Name]

            for e in accessory_elements:
                R = Reference(e)
                loc = e.Location.Point
                IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, loc)

            t.Commit()
        except Exception as e:
            t.RollBack()
            raise Exception("Failed to tag sleeves: {}".format(str(e)))

        tg.Assimilate()
    except Exception as e:
        tg.RollBack()
        raise
except Exception as e:
    print("Error: {}".format(str(e)))
    raise