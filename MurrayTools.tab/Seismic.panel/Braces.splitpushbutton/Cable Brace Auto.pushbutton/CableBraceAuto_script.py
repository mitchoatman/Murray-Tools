from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, Structure, XYZ, FabricationPart, FabricationConfiguration, TransactionGroup, BuiltInParameter, ElementTransformUtils, Line
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString
import math
import os
import sys

# Import for dialog
from rpw.ui.forms import FlexForm, Label, TextBox, Button

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

path, filename = os.path.split(__file__)
NewFilename = r'\CABLE SEISMIC BRACE.rfa'

families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'CABLE SEISMIC BRACE'
FamilyType = 'CABLE SEISMIC BRACE'

target_family = None
for f in families:
    if f.Name == FamilyName:
        target_family = f
        break

Fam_is_in_project = target_family is not None

def stretch_brace(family_instance, valuenum):
    set_parameter_by_name(family_instance, "Top of Steel", valuenum)
    BraceAngle = get_parameter_value_by_name(family_instance, "BraceMainAngle")
    sinofangle = math.sin(BraceAngle)
    BraceElevation = get_parameter_value_by_name(family_instance, 'Offset from Host')
    Height = ((valuenum - BraceElevation) - 0.2330)
    newhypotenus = ((Height / sinofangle) - 0.2290)
    if newhypotenus < 0:
        newhypotenus = 1
    set_parameter_by_name(family_instance, "BraceLength", (newhypotenus + 0.083333))

def set_parameter_by_name(element, parameterName, value):
    try:
        param = element.LookupParameter(parameterName)
        if param and not param.IsReadOnly:
            param.Set(value)
    except Exception:
        pass

def get_parameter_value_by_name(element, parameterName):
    try:
        param = element.LookupParameter(parameterName)
        if param:
            return param.AsDouble()
        return 0.0
    except Exception:
        return 0.0

def calculate_distance(point1, point2):
    return math.sqrt((point2.X - point1.X)**2 + (point2.Y - point1.Y)**2)

def calculate_angle(point1, point2):
    dx = point2.X - point1.X
    dy = point2.Y - point1.Y
    angle = math.atan2(dy, dx)
    return angle if angle >= 0 else angle + 2 * math.pi

def project_point_on_line(start_point, direction, point):
    v = XYZ(point.X - start_point.X, point.Y - start_point.Y, 0)
    direction = direction.Normalize()
    distance_along_line = v.DotProduct(direction)
    return distance_along_line

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie, service_name=None):
        self.nom_categorie = nom_categorie
        self.service_name = service_name
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            if self.service_name:
                return e.LookupParameter('Fabrication Service').AsValueString() == self.service_name
            return True
        return False
    def AllowReference(self, ref, point):
        return True

# Parse spacing input (feet, feet-inches, decimal)
def parse_spacing(input_str):
    try:
        input_str = input_str.strip()
        if '-' in input_str:  # e.g., "20-6"
            feet, inches = input_str.split('-')
            feet = float(feet.strip("'"))
            inches = float(inches.strip('"')) / 12.0
            return feet + inches
        elif "'" in input_str or '"' in input_str:  # e.g., "20'-0""
            input_str = input_str.replace("'", "").replace('"', "")
            parts = input_str.split('-')
            feet = float(parts[0])
            inches = float(parts[1]) / 12.0 if len(parts) > 1 else 0.0
            return feet + inches
        else:  # e.g., "20.5" or "20"
            return float(input_str)
    except (ValueError, IndexError):
        return None

# Prompt user for spacing
components = [
    Label('Transverse Spacing:'),
    TextBox('transverse_spacing', '20'),
    Label('Longitudinal Spacing:'),
    TextBox('longitudinal_spacing', '40'),
    Button('Ok')
]
form = FlexForm("Seismic Brace Spacing", components)
if not form.show():
    sys.exit(0)

transverse_spacing = parse_spacing(form.values["transverse_spacing"])
longitudinal_spacing = parse_spacing(form.values["longitudinal_spacing"])

# Validate inputs
if (transverse_spacing is None or longitudinal_spacing is None or
    transverse_spacing <= 0 or longitudinal_spacing <= 0):
    from rpw import ui
    ui.forms.Alert("Invalid spacing values. Please enter positive numbers.", exit=True)

try:
    first_hanger_sel = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter("MEP Fabrication Hangers"), "Select first hanger for initial brace")
    if first_hanger_sel is None:
        sys.exit(0)
except:
    sys.exit(0)

first_hanger = doc.GetElement(first_hanger_sel.ElementId)
first_hanger_service = first_hanger.LookupParameter('Fabrication Service').AsValueString()

try:
    pipesel = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter("MEP Fabrication Hangers", first_hanger_service), "Select multiple hangers to place seismic braces")
    if not pipesel or len(pipesel) == 0:
        sys.exit(0)
except:
    sys.exit(0)

Fhangers = [doc.GetElement(elId) for elId in pipesel]

family_pathCC = path + NewFilename

tg = TransactionGroup(doc, "Place Cable Brace Family")
tg.Start()

t = Transaction(doc, 'Load Cable Brace Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    target_family = doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()

familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralStiffener).OfClass(FamilySymbol).ToElements()
target_famtype = None

for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        target_famtype = famtype
        break

if target_famtype and len(Fhangers) > 1:
    t = Transaction(doc, 'Activate and Populate Cable Brace Family')
    t.Start()
    target_famtype.Activate()
    doc.Regenerate()

    hanger_positions = []
    for hanger in Fhangers:
        bounding_box = hanger.get_BoundingBox(None)
        if bounding_box:
            middle_top_point = XYZ((bounding_box.Min.X + bounding_box.Max.X) / 2,
                                 (bounding_box.Min.Y + bounding_box.Max.Y) / 2,
                                 bounding_box.Max.Z)
            middle_bottom_point = XYZ((bounding_box.Min.X + bounding_box.Max.X) / 2,
                                    (bounding_box.Min.Y + bounding_box.Max.Y) / 2,
                                    bounding_box.Min.Z)
            hanger_positions.append((hanger, middle_top_point, middle_bottom_point))

    first_hanger_id = first_hanger.Id
    first_hanger_index = next((i for i, hp in enumerate(hanger_positions) if hp[0].Id == first_hanger_id), 0)
    first_point = hanger_positions[first_hanger_index][1]

    max_distance = 0
    farthest_index = 0
    for i, (hanger, point, _) in enumerate(hanger_positions):
        dist = calculate_distance(first_point, point)
        if dist > max_distance:
            max_distance = dist
            farthest_index = i
    
    farthest_point = hanger_positions[farthest_index][1]
    
    direction = XYZ(farthest_point.X - first_point.X, farthest_point.Y - first_point.Y, 0)
    base_angle = calculate_angle(first_point, farthest_point)

    hanger_positions.sort(key=lambda x: project_point_on_line(first_point, direction, x[1]))

    braces_placed = set()
    last_brace_position = hanger_positions[0][1]
    transverse_hanger_indices = [0]  # Track indices of hangers with transverse braces

    # First hanger braces (transverse, doubled up)
    hanger = hanger_positions[0][0]
    middle_top = hanger_positions[0][1]
    middle_bottom = hanger_positions[0][2]
    STName = hanger.GetRodInfo().RodCount
    STName1 = hanger.GetRodInfo()
    RackType = get_parameter_value_by_name_AsValueString(hanger, 'Family')

    ItmDims = hanger.GetDimensions()
    BraceOffsetZ = 0.0
    for dta in ItmDims:
        if dta.Name in ['Rod Length', 'RodLength', 'Rod Extn Above']:
            BraceOffsetZ = hanger.GetDimensionValue(dta)
            break
    if BraceOffsetZ == 0.0:
        BraceOffsetZ = 0.0

    if STName == 1:
        rodloc = STName1.GetRodEndPosition(0)
        valuenum = rodloc.Z
        pos_key = (round(middle_top.X, 4), round(middle_top.Y, 4))
        pos_key2 = (round(middle_top.X, 4), round(middle_top.Y, 4), 2)
        new_insertion_point = XYZ(middle_top.X, middle_top.Y, middle_top.Z + 0.025716145 - BraceOffsetZ)
        if pos_key not in braces_placed:
            new_family_instance = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
            ElementTransformUtils.RotateElement(doc, new_family_instance.Id, Line.CreateBound(new_insertion_point, XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)), base_angle + math.pi/2)
            stretch_brace(new_family_instance, valuenum)
            braces_placed.add(pos_key)
        if pos_key2 not in braces_placed:
            new_family_instance2 = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
            ElementTransformUtils.RotateElement(doc, new_family_instance2.Id, Line.CreateBound(new_insertion_point, XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)), base_angle + math.pi + math.pi/2)
            stretch_brace(new_family_instance2, valuenum)
            braces_placed.add(pos_key2)
    else:
        for n in range(STName):
            rodloc = STName1.GetRodEndPosition(n)
            valuenum = rodloc.Z
            pos_key = (round(rodloc.X, 4), round(rodloc.Y, 4), n)
            if RackType == '1.625 Single Strut Trapeze':
                combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.25524)
            elif RackType == '038 Unistrut Trapeeze':
                combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.25)
            elif RackType == '050 Unistrut Trapeeze':
                combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.2815)
            elif RackType == '1.625 Double Strut Trapeze':
                combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.390656)
            elif RackType == '050 Doublestrut Trapeeze':
                combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.4165)
            elif 'Seismic' in RackType:
                combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_top.Z - BraceOffsetZ + 0.13541)
            else:
                combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z)
            rodloc1 = STName1.GetRodEndPosition(0)
            rodloc2 = STName1.GetRodEndPosition(1)
            rod_angle = calculate_angle(rodloc2, rodloc1)
            brace_angle = rod_angle + math.pi if n == 0 else rod_angle
            if pos_key not in braces_placed:
                new_family_instance = doc.Create.NewFamilyInstance(combined_xyz, target_famtype, DB.Structure.StructuralType.NonStructural)
                ElementTransformUtils.RotateElement(doc, new_family_instance.Id, Line.CreateBound(combined_xyz, XYZ(combined_xyz.X, combined_xyz.Y, combined_xyz.Z + 1)), brace_angle)
                stretch_brace(new_family_instance, valuenum)
                braces_placed.add(pos_key)

    # Place transverse braces
    reference_position = hanger_positions[0][1]
    i = 1
    while i < len(hanger_positions):
        max_distance = -1
        closest_hanger_index = -1
        for j in range(i, len(hanger_positions)):
            total_distance = calculate_distance(reference_position, hanger_positions[j][1])
            if total_distance < transverse_spacing and total_distance > max_distance:
                max_distance = total_distance
                closest_hanger_index = j
        
        if closest_hanger_index != -1:
            transverse_hanger_indices.append(closest_hanger_index)
            hanger = hanger_positions[closest_hanger_index][0]
            middle_top = hanger_positions[closest_hanger_index][1]
            middle_bottom = hanger_positions[closest_hanger_index][2]
            STName = hanger.GetRodInfo().RodCount
            STName1 = hanger.GetRodInfo()
            RackType = get_parameter_value_by_name_AsValueString(hanger, 'Family')
            
            ItmDims = hanger.GetDimensions()
            BraceOffsetZ = 0.0
            for dta in ItmDims:
                if dta.Name in ['Rod Length', 'RodLength', 'Rod Extn Above']:
                    BraceOffsetZ = hanger.GetDimensionValue(dta)
                    break
            if BraceOffsetZ == 0.0:
                BraceOffsetZ = 0.0
            
            if STName == 1:
                rodloc = STName1.GetRodEndPosition(0)
                valuenum = rodloc.Z
                pos_key = (round(middle_top.X, 4), round(middle_top.Y, 4))
                pos_key2 = (round(middle_top.X, 4), round(middle_top.Y, 4), 2)
                new_insertion_point = XYZ(middle_top.X, middle_top.Y, middle_top.Z + 0.025716145 - BraceOffsetZ)
                if pos_key not in braces_placed:
                    new_family_instance = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
                    ElementTransformUtils.RotateElement(doc, new_family_instance.Id, Line.CreateBound(new_insertion_point, XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)), base_angle + math.pi/2)
                    stretch_brace(new_family_instance, valuenum)
                    braces_placed.add(pos_key)
                if pos_key2 not in braces_placed:
                    new_family_instance2 = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
                    ElementTransformUtils.RotateElement(doc, new_family_instance2.Id, Line.CreateBound(new_insertion_point, XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)), base_angle + math.pi + math.pi/2)
                    stretch_brace(new_family_instance2, valuenum)
                    braces_placed.add(pos_key2)
            else:
                for n in range(STName):
                    rodloc = STName1.GetRodEndPosition(n)
                    valuenum = rodloc.Z
                    pos_key = (round(rodloc.X, 4), round(rodloc.Y, 4), n)
                    if RackType == '1.625 Single Strut Trapeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.25524)
                    elif RackType == '038 Unistrut Trapeeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.25)
                    elif RackType == '050 Unistrut Trapeeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.2815)
                    elif RackType == '1.625 Double Strut Trapeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.390656)
                    elif RackType == '050 Doublestrut Trapeeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.4165)
                    elif 'Seismic' in RackType:
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_top.Z - BraceOffsetZ + 0.13541)
                    else:
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z)
                    rodloc1 = STName1.GetRodEndPosition(0)
                    rodloc2 = STName1.GetRodEndPosition(1)
                    rod_angle = calculate_angle(rodloc2, rodloc1)
                    brace_angle = rod_angle + math.pi if n == 0 else rod_angle
                    if pos_key not in braces_placed:
                        new_family_instance = doc.Create.NewFamilyInstance(combined_xyz, target_famtype, DB.Structure.StructuralType.NonStructural)
                        ElementTransformUtils.RotateElement(doc, new_family_instance.Id, Line.CreateBound(combined_xyz, XYZ(combined_xyz.X, combined_xyz.Y, combined_xyz.Z + 1)), brace_angle)
                        stretch_brace(new_family_instance, valuenum)
                        braces_placed.add(pos_key)
            
            reference_position = middle_top
            last_brace_position = middle_top
            i = closest_hanger_index + 1
        else:
            break

    # Place longitudinal braces, excluding last hanger
    reference_position = hanger_positions[0][1]
    current_transverse_index = 0
    last_hanger_index = len(hanger_positions) - 1
    while current_transverse_index < len(transverse_hanger_indices):
        max_distance = -1
        closest_hanger_index = -1
        for j in range(current_transverse_index, len(transverse_hanger_indices)):
            hanger_index = transverse_hanger_indices[j]
            total_distance = calculate_distance(reference_position, hanger_positions[hanger_index][1])
            if total_distance < longitudinal_spacing and total_distance > max_distance and hanger_index != last_hanger_index:
                max_distance = total_distance
                closest_hanger_index = hanger_index
        
        if closest_hanger_index != -1:
            hanger = hanger_positions[closest_hanger_index][0]
            middle_top = hanger_positions[closest_hanger_index][1]
            middle_bottom = hanger_positions[closest_hanger_index][2]
            STName = hanger.GetRodInfo().RodCount
            STName1 = hanger.GetRodInfo()
            RackType = get_parameter_value_by_name_AsValueString(hanger, 'Family')
            
            ItmDims = hanger.GetDimensions()
            BraceOffsetZ = 0.0
            for dta in ItmDims:
                if dta.Name in ['Rod Length', 'RodLength', 'Rod Extn Above']:
                    BraceOffsetZ = hanger.GetDimensionValue(dta)
                    break
            if BraceOffsetZ == 0.0:
                BraceOffsetZ = 0.0
            
            if STName == 1:
                rodloc = STName1.GetRodEndPosition(0)
                valuenum = rodloc.Z
                pos_key3 = (round(middle_top.X, 4), round(middle_top.Y, 4), 3)
                pos_key4 = (round(middle_top.X, 4), round(middle_top.Y, 4), 4)
                new_insertion_point = XYZ(middle_top.X, middle_top.Y, middle_top.Z + 0.025716145 - BraceOffsetZ)
                if pos_key3 not in braces_placed:
                    new_family_instance3 = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
                    ElementTransformUtils.RotateElement(doc, new_family_instance3.Id, Line.CreateBound(new_insertion_point, XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)), base_angle)
                    stretch_brace(new_family_instance3, valuenum)
                    braces_placed.add(pos_key3)
                if pos_key4 not in braces_placed:
                    new_family_instance4 = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
                    ElementTransformUtils.RotateElement(doc, new_family_instance4.Id, Line.CreateBound(new_insertion_point, XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)), base_angle + math.pi)
                    stretch_brace(new_family_instance4, valuenum)
                    braces_placed.add(pos_key4)
            else:
                for n in range(STName):
                    rodloc = STName1.GetRodEndPosition(n)
                    valuenum = rodloc.Z
                    pos_key1 = (round(rodloc.X, 4), round(rodloc.Y, 4), n * 2 + 3)
                    pos_key2 = (round(rodloc.X, 4), round(rodloc.Y, 4), n * 2 + 4)
                    if RackType == '1.625 Single Strut Trapeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.25524)
                    elif RackType == '038 Unistrut Trapeeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.25)
                    elif RackType == '050 Unistrut Trapeeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.2815)
                    elif RackType == '1.625 Double Strut Trapeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.390656)
                    elif RackType == '050 Doublestrut Trapeeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z + 0.4165)
                    elif 'Seismic' in RackType:
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_top.Z - BraceOffsetZ + 0.13541)
                    else:
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom.Z)
                    if pos_key1 not in braces_placed:
                        new_family_instance1 = doc.Create.NewFamilyInstance(combined_xyz, target_famtype, DB.Structure.StructuralType.NonStructural)
                        ElementTransformUtils.RotateElement(doc, new_family_instance1.Id, Line.CreateBound(combined_xyz, XYZ(combined_xyz.X, combined_xyz.Y, combined_xyz.Z + 1)), base_angle)
                        stretch_brace(new_family_instance1, valuenum)
                        braces_placed.add(pos_key1)
                    if pos_key2 not in braces_placed:
                        new_family_instance2 = doc.Create.NewFamilyInstance(combined_xyz, target_famtype, DB.Structure.StructuralType.NonStructural)
                        ElementTransformUtils.RotateElement(doc, new_family_instance2.Id, Line.CreateBound(combined_xyz, XYZ(combined_xyz.X, combined_xyz.Y, combined_xyz.Z + 1)), base_angle + math.pi)
                        stretch_brace(new_family_instance2, valuenum)
                        braces_placed.add(pos_key2)
            
            reference_position = middle_top
            last_brace_position = middle_top
            for idx, hanger_idx in enumerate(transverse_hanger_indices):
                if hanger_idx == closest_hanger_index:
                    current_transverse_index = idx + 1
                    break
        else:
            break

    # Last hanger logic (transverse only if > transverse_spacing/2)
    if len(hanger_positions) > 1:
        last_hanger = hanger_positions[-1][0]
        last_middle_top = hanger_positions[-1][1]
        last_middle_bottom = hanger_positions[-1][2]
        last_distance = calculate_distance(last_brace_position, last_middle_top)
        if last_distance > transverse_spacing / 2.0:
            STName = last_hanger.GetRodInfo().RodCount
            STName1 = last_hanger.GetRodInfo()
            RackType = get_parameter_value_by_name_AsValueString(last_hanger, 'Family')
            
            ItmDims = last_hanger.GetDimensions()
            BraceOffsetZ = 0.0
            for dta in ItmDims:
                if dta.Name in ['Rod Length', 'RodLength', 'Rod Extn Above']:
                    BraceOffsetZ = last_hanger.GetDimensionValue(dta)
                    break
            if BraceOffsetZ == 0.0:
                BraceOffsetZ = 0.0
            
            if STName == 1:
                rodloc = STName1.GetRodEndPosition(0)
                valuenum = rodloc.Z
                pos_key_20 = (round(last_middle_top.X, 4), round(last_middle_top.Y, 4))
                pos_key_20_2 = (round(last_middle_top.X, 4), round(last_middle_top.Y, 4), 2)
                new_insertion_point = XYZ(last_middle_top.X, last_middle_top.Y, last_middle_top.Z + 0.025716145 - BraceOffsetZ)
                if pos_key_20 not in braces_placed:
                    new_family_instance = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
                    ElementTransformUtils.RotateElement(doc, new_family_instance.Id, Line.CreateBound(new_insertion_point, XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)), base_angle + math.pi/2)
                    stretch_brace(new_family_instance, valuenum)
                    braces_placed.add(pos_key_20)
                if pos_key_20_2 not in braces_placed:
                    new_family_instance2 = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
                    ElementTransformUtils.RotateElement(doc, new_family_instance2.Id, Line.CreateBound(new_insertion_point, XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)), base_angle + math.pi + math.pi/2)
                    stretch_brace(new_family_instance2, valuenum)
                    braces_placed.add(pos_key_20_2)
            else:
                for n in range(STName):
                    rodloc = STName1.GetRodEndPosition(n)
                    valuenum = rodloc.Z
                    pos_key = (round(rodloc.X, 4), round(rodloc.Y, 4), n)
                    if RackType == '1.625 Single Strut Trapeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, last_middle_bottom.Z + 0.25524)
                    elif RackType == '038 Unistrut Trapeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, last_middle_bottom.Z + 0.25)
                    elif RackType == '050 Unistrut Trapeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, last_middle_bottom.Z + 0.2815)
                    elif RackType == '1.625 Double Strut Trapeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, last_middle_bottom.Z + 0.390656)
                    elif RackType == '050 Doublestrut Trapeze':
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, last_middle_bottom.Z + 0.4165)
                    elif 'Seismic' in RackType:
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, last_middle_top.Z - BraceOffsetZ + 0.13541)
                    else:
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, last_middle_bottom.Z)
                    rodloc1 = STName1.GetRodEndPosition(0)
                    rodloc2 = STName1.GetRodEndPosition(1)
                    rod_angle = calculate_angle(rodloc2, rodloc1)
                    brace_angle = rod_angle + math.pi if n == 0 else rod_angle
                    if pos_key not in braces_placed:
                        new_family_instance = doc.Create.NewFamilyInstance(combined_xyz, target_famtype, DB.Structure.StructuralType.NonStructural)
                        ElementTransformUtils.RotateElement(doc, new_family_instance.Id, Line.CreateBound(combined_xyz, XYZ(combined_xyz.X, combined_xyz.Y, combined_xyz.Z + 1)), brace_angle)
                        stretch_brace(new_family_instance, valuenum)
                        braces_placed.add(pos_key)

    t.Commit()
tg.Assimilate()