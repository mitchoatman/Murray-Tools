# -*- coding: utf-8 -*-
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, Structure, XYZ, TransactionGroup, FamilyInstance
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import os, sys

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

# Family loading options
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

# Custom selection filter
class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        return e.Category and e.Category.Name == self.nom_categorie
    def AllowReference(self, ref, point):
        return True

# Compute top center of bounding box
def get_bbox_top_center(element):
    bbox = element.get_BoundingBox(None)
    if not bbox:
        return None
    return XYZ(
        (bbox.Min.X + bbox.Max.X) / 2,
        (bbox.Min.Y + bbox.Max.Y) / 2,
        bbox.Max.Z
    )

try:
    # Prompt user to select one pipe accessory
    pipesel = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter("Pipe Accessories"), "Select Fabrication Hangers to place seismic brace on")
    if not pipesel:
        sys.exit(0)  # Exit gracefully if user cancels

    pipe_acc = doc.GetElement(pipesel.ElementId)

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
            # Activate family symbol
            if not target_famtype.IsActive:
                t = Transaction(doc, "Activate Family Symbol")
                t.Start()
                target_famtype.Activate()
                doc.Regenerate()
                t.Commit()

            # Determine placement targets
            parent_bbox = pipe_acc.get_BoundingBox(None)
            all_fis = FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements()
            nested_fams = []
            for fi in all_fis:
                # Only generic models
                if fi.Symbol.Family.FamilyCategory.Name != "Generic Models":
                    continue
                fi_bbox = fi.get_BoundingBox(None)
                if not fi_bbox:
                    continue
                # Check XY containment inside parent
                if (parent_bbox.Min.X <= fi_bbox.Min.X <= parent_bbox.Max.X and
                    parent_bbox.Min.Y <= fi_bbox.Min.Y <= parent_bbox.Max.Y):
                    nested_fams.append(fi)

            placement_targets = nested_fams if nested_fams else [pipe_acc]

            # Place braces in one transaction
            t = Transaction(doc, 'Place Seismic Brace')
            t.Start()
            level = doc.ActiveView.GenLevel  # Use view level for placement
            for elem in placement_targets:
                top_center = get_bbox_top_center(elem)
                if not top_center:
                    continue  # skip if invalid
                try:
                    doc.Create.NewFamilyInstance(
                        top_center,
                        target_famtype,
                        level,
                        Structure.StructuralType.NonStructural
                    )
                except Exception as e:
                    print("Failed to place brace on element ID:", elem.Id)
            t.Commit()

        tg.Assimilate()

except Autodesk.Revit.Exceptions.OperationCanceledException:
    sys.exit(0)
