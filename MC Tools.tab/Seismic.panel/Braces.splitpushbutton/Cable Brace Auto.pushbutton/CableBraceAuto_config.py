# coding: utf8
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family,
    Structure, XYZ, TransactionGroup, BuiltInParameter,
    Line, ElementTransformUtils
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Parameters.Get_Set_Params import (
    get_parameter_value_by_name_AsValueString,
    get_parameter_value_by_name_AsDouble,
    set_parameter_by_name,
    get_parameter_value_by_name_AsString
)
from Autodesk.Revit.UI import TaskDialog
import math
import os
import sys

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

path, filename = os.path.split(__file__)
NewFilename = r'\CABLE SEISMIC BRACE.rfa'

# ---------------- FAMILY ----------------
families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'CABLE SEISMIC BRACE'
FamilyType = 'CABLE SEISMIC BRACE'

target_family = None
for f in families:
    if f.Name == FamilyName:
        target_family = f
        break

Fam_is_in_project = target_family is not None

# ---------------- STRETCH ----------------
def stretch_brace():
    try:
        set_parameter_by_name(new_family_instance, "Top of Steel", valuenum)
        BraceAngle = get_parameter_value_by_name_AsDouble(new_family_instance, "BraceMainAngle")
        sinofangle = math.sin(BraceAngle)
        BraceElevation = get_parameter_value_by_name_AsDouble(new_family_instance, 'Offset from Host')

        Height = ((valuenum - BraceElevation) - 0.2330)
        newhypotenus = ((Height / sinofangle) - 0.175)

        if newhypotenus < 0:
            newhypotenus = 1

        set_parameter_by_name(new_family_instance, "BraceLength", newhypotenus)
        set_parameter_by_name(new_family_instance, "ISAT Brace Level", HangerLevel)

        if new_family_instance.LookupParameter("FP_Service Name"):
            set_parameter_by_name(new_family_instance, "FP_Service Name", HangerService)
    except:
        pass

# ---------------- 4 BRACE PLACEMENT ----------------
def place_four_braces(base_point, offset_point=None):
    global new_family_instance

    if offset_point is None:
        offset_point = base_point

    axis = Line.CreateBound(base_point, XYZ(base_point.X, base_point.Y, base_point.Z + 1))

    # 0°
    new_family_instance = doc.Create.NewFamilyInstance(base_point, target_famtype, DB.Structure.StructuralType.NonStructural)
    stretch_brace()

    # 90°
    inst = doc.Create.NewFamilyInstance(offset_point, target_famtype, DB.Structure.StructuralType.NonStructural)
    ElementTransformUtils.RotateElement(doc, inst.Id, axis, math.pi / 2)
    new_family_instance = inst
    stretch_brace()

    # 180°
    inst = doc.Create.NewFamilyInstance(base_point, target_famtype, DB.Structure.StructuralType.NonStructural)
    ElementTransformUtils.RotateElement(doc, inst.Id, axis, math.pi)
    new_family_instance = inst
    stretch_brace()

    # 270°
    inst = doc.Create.NewFamilyInstance(offset_point, target_famtype, DB.Structure.StructuralType.NonStructural)
    ElementTransformUtils.RotateElement(doc, inst.Id, axis, 3 * math.pi / 2)
    new_family_instance = inst
    stretch_brace()

# ---------------- SELECTION ----------------
class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, name):
        self.name = name
    def AllowElement(self, e):
        return e.Category.Name == self.name
    def AllowReference(self, ref, point):
        return True

try:
    sel = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter("MEP Fabrication Hangers"), "Select Hangers")
except:
    sys.exit()

Fhangers = [doc.GetElement(x) for x in sel]

# ---------------- LOAD FAMILY ----------------
tg = TransactionGroup(doc, "Cable Brace 4-Way")
tg.Start()

t = Transaction(doc, "Load Family")
t.Start()
if not Fam_is_in_project:
    doc.LoadFamily(path + NewFilename)
t.Commit()

# ---------------- GET TYPE ----------------
types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralStiffener).OfClass(FamilySymbol)

target_famtype = None
for ttype in types:
    if ttype.Family.Name == FamilyName and ttype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType:
        target_famtype = ttype
        break

# ---------------- PLACE ----------------
if target_famtype:
    t = Transaction(doc, "Place 4 Braces")
    t.Start()

    target_famtype.Activate()
    doc.Regenerate()

    for hanger in Fhangers:

        RackType = get_parameter_value_by_name_AsValueString(hanger, 'Family')
        if 'strap' in RackType.lower():
            TaskDialog.Show("Error", "Cannot Place Seismic on that Support!")
            continue

        bbox = hanger.get_BoundingBox(None)
        if not bbox:
            continue

        mid_top = XYZ((bbox.Min.X + bbox.Max.X)/2, (bbox.Min.Y + bbox.Max.Y)/2, bbox.Max.Z)
        mid_bot = XYZ((bbox.Min.X + bbox.Max.X)/2, (bbox.Min.Y + bbox.Max.Y)/2, bbox.Min.Z)

        HangerLevel = get_parameter_value_by_name_AsValueString(hanger, 'Reference Level')
        HangerService = get_parameter_value_by_name_AsString(hanger, 'Fabrication Service Name')

        rodinfo = hanger.GetRodInfo()
        count = rodinfo.RodCount

        BraceOffsetZ = 0
        for d in hanger.GetDimensions():
            if d.Name in ['Rod Length', 'RodLength', 'Rod Extn Above']:
                BraceOffsetZ = hanger.GetDimensionValue(d)
                break

        if count == 1:
            rod = rodinfo.GetRodEndPosition(0)
            valuenum = rod.Z

            base = XYZ(mid_top.X, mid_top.Y, mid_top.Z + 0.025716145 - BraceOffsetZ)
            offset = XYZ(base.X, base.Y, base.Z + 0.03125)

            place_four_braces(base, offset)

        else:
            for i in range(count):
                rod = rodinfo.GetRodEndPosition(i)
                valuenum = rod.Z

                if "1.625" in RackType.lower():
                    z1 = mid_bot.Z + 0.25524
                    z2 = mid_bot.Z + 0.28649
                elif "unistrut" in RackType.lower():
                    z1 = mid_bot.Z + 0.25
                    z2 = mid_bot.Z + 0.28125
                elif "seismic" in RackType.lower():
                    z1 = mid_top.Z - BraceOffsetZ + 0.1615
                    z2 = mid_top.Z - BraceOffsetZ + 0.19275
                else:
                    z1 = mid_top.Z - BraceOffsetZ
                    z2 = z1 + 0.03125

                base = XYZ(rod.X, rod.Y, z1)
                offset = XYZ(rod.X, rod.Y, z2)

                place_four_braces(base, offset)

    t.Commit()

tg.Assimilate()