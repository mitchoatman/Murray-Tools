import Autodesk
from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource, Transaction, FilteredElementCollector, Family, TransactionGroup,\
                                BuiltInCategory, FamilySymbol, BuiltInParameter, Reference, IndependentTag, TagMode, TagOrientation
import os
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialog

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

try:
    # Search project for all Families
    families = FilteredElementCollector(doc).OfClass(Family)
    FamilyName = 'Multi Category Tag - FP_Valve Number'
    FamilyType = 'Multi Category Tag - FP_Valve Number'
    Fam_is_in_project = any(f.Name == FamilyName for f in families)

    class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
        def OnFamilyFound(self, familyInUse, overwriteParameterValues):
            overwriteParameterValues.Value = False
            return True

        def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
            source.Value = DB.FamilySource.Family
            overwriteParameterValues.Value = False
            return True

    family_path = r'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES\Annotation\Multi Category Tag - FP_Valve Number.rfa'

    tg = TransactionGroup(doc, "Add Valve Tags")
    try:
        tg.Start()

        t = Transaction(doc, 'Load Valve Tag Family')
        try:
            t.Start()
            if not Fam_is_in_project:
                fload_handler = FamilyLoaderOptionsHandler()
                if not os.path.exists(family_path):
                    raise Exception("Family file not found: {}".format(family_path))
                family = doc.LoadFamily(family_path, fload_handler)
            t.Commit()
        except Exception as e:
            t.RollBack()
            raise Exception("Failed to load family: {}".format(str(e)))

        try:
            Pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework)\
                                                                    .WhereElementIsNotElementType()
            familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation)\
                                                                .OfClass(FamilySymbol)\
                                                                .ToElements()
            existing_tags = FilteredElementCollector(doc, curview.Id).OfClass(IndependentTag).ToElements()
        except Exception as e:
            raise Exception("Failed to collect elements: {}".format(str(e)))

        t = Transaction(doc, 'Tag Valves')
        try:
            t.Start()
            for famtype in familyTypes:
                typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                if famtype.Family.Name == FamilyName and typeName == FamilyType:
                    if not famtype.IsActive:
                        famtype.Activate()
                        doc.Regenerate()

            for valve in Pipe_collector:
                ST = valve.ServiceType
                AL = valve.Alias
                if ST == 53 and AL not in ['STRAINER', 'CHECK', 'BALANCE']:
                    is_tagged = False
                    valve_id = valve.Id
                    for tag in existing_tags:
                        tagged_ids = tag.GetTaggedLocalElementIds()
                        if valve_id in tagged_ids:
                            is_tagged = True
                            break
                    
                    if not is_tagged:
                        R = Reference(valve)
                        ValveLocation = valve.Origin
                        IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_MULTICATEGORY, TagOrientation.Horizontal, ValveLocation)

            t.Commit()
        except Exception as e:
            t.RollBack()
            TaskDialog.Show("Error", "Failed to tag valves: {}".format(str(e)))
            # raise Exception("Failed to tag valves: {}".format(str(e)))

        tg.Assimilate()
    except Exception as e:
        tg.RollBack()
        raise
except Exception as e:
    TaskDialog.Show("Error", "Failed to tag valves: {}".format(str(e)))
    raise