from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, Structure, XYZ, TransactionGroup
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import os

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

path, filename = os.path.split(__file__)
NewFilename = r'\RIGID SEISMIC BRACE.rfa'

# Family setup
FamilyName = 'RIGID SEISMIC BRACE'
FamilyType = 'RIGID SEISMIC BRACE'

# Check if family is in project
families = FilteredElementCollector(doc).OfClass(Family)
target_family = next((f for f in families if f.Name == FamilyName), None)
Fam_is_in_project = target_family is not None

# Selection filter for pipe accessories
class PipeAccessorySelectionFilter(ISelectionFilter):
    def AllowElement(self, e):
        return e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeAccessory)
    def AllowReference(self, ref, point):
        return True

# Family loading options
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

try:
    # Prompt user to select one pipe accessory
    pipe_acc_ref = uidoc.Selection.PickObject(ObjectType.Element, 
                                            PipeAccessorySelectionFilter(), 
                                            "Select a Pipe Accessory for seismic brace placement")
    pipe_acc = doc.GetElement(pipe_acc_ref.ElementId)

    # Start transaction group
    with TransactionGroup(doc, "Place Seismic Brace") as tg:
        tg.Start()

        # Load family if not present
        if not Fam_is_in_project:
            t = Transaction(doc, 'Load Rigid Brace Family')
            t.Start()
            fload_handler = FamilyLoaderOptionsHandler()
            family_path = path + NewFilename
            target_family = doc.LoadFamily(family_path, fload_handler)
            t.Commit()

        # Get family symbol
        family_symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralStiffener).OfClass(FamilySymbol)
        target_famtype = next((fs for fs in family_symbols 
                             if fs.Family.Name == FamilyName and 
                             fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType), None)

        if target_famtype:
            t = Transaction(doc, 'Place Seismic Brace')
            t.Start()
            
            # Activate symbol
            target_famtype.Activate()
            doc.Regenerate()

            # Get bounding box and calculate center/top point
            bbox = pipe_acc.get_BoundingBox(None)
            if bbox:
                center_top_point = XYZ(
                    (bbox.Min.X + bbox.Max.X) / 2,
                    (bbox.Min.Y + bbox.Max.Y) / 2,
                    bbox.Max.Z
                )
                
                # Place the brace
                new_brace = doc.Create.NewFamilyInstance(
                    center_top_point,
                    target_famtype,
                    DB.Structure.StructuralType.NonStructural
                )

            t.Commit()
        tg.Assimilate()

except Exception as e:
    print("Error: {}".format(str(e)))