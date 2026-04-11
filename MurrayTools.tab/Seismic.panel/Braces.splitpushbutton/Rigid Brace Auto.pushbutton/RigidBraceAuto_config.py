# coding: utf8
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family,
    Structure, XYZ, FabricationPart, FabricationConfiguration, TransactionGroup,
    BuiltInParameter, ElementTransformUtils, Line
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString, get_parameter_value_by_name_AsDouble, set_parameter_by_name, get_parameter_value_by_name_AsString
from Parameters.Add_SharedParameters import Shared_Params
from Autodesk.Revit.UI import TaskDialog
import math
import os
import sys

Shared_Params()

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

path, filename = os.path.split(__file__)
NewFilename = r'\RIGID SEISMIC BRACE.rfa'

families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'RIGID SEISMIC BRACE'
FamilyType = 'RIGID SEISMIC BRACE'

target_family = None
for f in families:
    if f.Name == FamilyName:
        target_family = f
        break

Fam_is_in_project = target_family is not None

def stretch_brace():
    try:
        set_parameter_by_name(new_family_instance, "Top of Steel", valuenum)
        BraceAngle = get_parameter_value_by_name_AsDouble(new_family_instance, "BraceMainAngle")
        sinofangle = math.sin(BraceAngle)
        BraceElevation = get_parameter_value_by_name_AsDouble(new_family_instance, 'Offset from Host')
        Height = ((valuenum - BraceElevation) - 0.2330)
        newhypotenus = ((Height / sinofangle) - 0.2290)
        if newhypotenus < 0:
            newhypotenus = 1
        set_parameter_by_name(new_family_instance, "BraceLength", newhypotenus)
        set_parameter_by_name(new_family_instance, "ISAT Brace Level", HangerLevel)
        if new_family_instance.LookupParameter("FP_Service Name"):
            set_parameter_by_name(new_family_instance, "FP_Service Name", HangerService)
    except:
        pass

def place_two_braces(point):
    global new_family_instance

    # First brace
    new_family_instance = doc.Create.NewFamilyInstance(point, target_famtype, DB.Structure.StructuralType.NonStructural)
    stretch_brace()

    # Rotation axis (vertical)
    axis_start = point
    axis_end = XYZ(point.X, point.Y, point.Z + 1)
    rotation_axis = Line.CreateBound(axis_start, axis_end)

    # Second brace point (Z + 3/8" = 0.03125 ft)
    offset_point = XYZ(point.X, point.Y, point.Z + 0.04)

    # Second brace
    second_instance = doc.Create.NewFamilyInstance(offset_point, target_famtype, DB.Structure.StructuralType.NonStructural)
    ElementTransformUtils.RotateElement(doc, second_instance.Id, rotation_axis, math.pi / 2)

    new_family_instance = second_instance
    stretch_brace()

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        return e.Category.Name == self.nom_categorie
    def AllowReference(self, ref, point):
        return True

try:
    pipesel = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")
    if not pipesel:
        sys.exit(0)
except Autodesk.Revit.Exceptions.OperationCanceledException:
    sys.exit(0)

Fhangers = [doc.GetElement(elId) for elId in pipesel]

family_pathCC = path + NewFilename

tg = TransactionGroup(doc, "Place Rigid Brace Family")
tg.Start()

t = Transaction(doc, 'Load Rigid Brace Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    target_family = doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()

familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralStiffener).OfClass(FamilySymbol)

target_famtype = None
for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        target_famtype = famtype
        break

if target_famtype:
    t = Transaction(doc, 'Place Braces')
    t.Start()

    target_famtype.Activate()
    doc.Regenerate()

    for hanger in Fhangers:

        RackType = get_parameter_value_by_name_AsValueString(hanger, 'Family')
        rack_type_lower = RackType.lower()

        if 'strap' in rack_type_lower:
            TaskDialog.Show("Error", "Cannot Place Seismic on that Support!")
            continue

        bbox = hanger.get_BoundingBox(None)
        if not bbox:
            continue

        middle_top = XYZ((bbox.Min.X + bbox.Max.X)/2, (bbox.Min.Y + bbox.Max.Y)/2, bbox.Max.Z)
        middle_bottom = XYZ((bbox.Min.X + bbox.Max.X)/2, (bbox.Min.Y + bbox.Max.Y)/2, bbox.Min.Z)

        HangerLevel = get_parameter_value_by_name_AsValueString(hanger, 'Reference Level')
        HangerService = get_parameter_value_by_name_AsString(hanger, 'Fabrication Service Name')

        rod_info = hanger.GetRodInfo()
        rod_count = rod_info.RodCount

        BraceOffsetZ = 0
        for d in hanger.GetDimensions():
            if d.Name in ['Rod Length', 'RodLength', 'Rod Extn Above']:
                BraceOffsetZ = hanger.GetDimensionValue(d)
                break

        if rod_count == 1:
            rodloc = rod_info.GetRodEndPosition(0)
            valuenum = rodloc.Z

            point = XYZ(middle_top.X, middle_top.Y, middle_top.Z - BraceOffsetZ)
            place_two_braces(point)

        else:
            for i in range(rod_count):
                rodloc = rod_info.GetRodEndPosition(i)
                valuenum = rodloc.Z

                if "1.625" in rack_type_lower:
                    z = middle_bottom.Z + (0.2292 if "single" in rack_type_lower else 0.3646)
                elif "unistrut" in rack_type_lower:
                    if "038" in rack_type_lower or "1-5/8" in rack_type_lower:
                        z = middle_bottom.Z + 0.25
                    elif "050" in rack_type_lower:
                        z = middle_bottom.Z + (0.4165 if "double" in rack_type_lower else 0.2815)
                elif "seismic" in rack_type_lower:
                    z = middle_top.Z - BraceOffsetZ + 0.13541
                else:
                    z = middle_top.Z - BraceOffsetZ

                point = XYZ(rodloc.X, rodloc.Y, z)
                place_two_braces(point)

    t.Commit()

tg.Assimilate()