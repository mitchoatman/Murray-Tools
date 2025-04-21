import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, Structure, XYZ, FabricationPart, FabricationConfiguration, TransactionGroup, BuiltInParameter, Line, ElementTransformUtils
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString
import math
import os
import sys  # Added for sys.exit

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

path, filename = os.path.split(__file__)
NewFilename = r'\CABLE SEISMIC BRACE.rfa'

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'CABLE SEISMIC BRACE'
FamilyType = 'CABLE SEISMIC BRACE'

# Retrieve the specific family by name
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
    element.LookupParameter(parameterName).Set(value)

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def get_parameter_value_by_name_AsDouble(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

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
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return True

# Handle user cancellation for hanger selection
try:
    pipesel = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers to place seismic brace on")
    if not pipesel:  # Check if selection is empty
        sys.exit(0)  # Exit gracefully if no hangers selected
except Autodesk.Revit.Exceptions.OperationCanceledException:
    sys.exit(0)  # Exit gracefully if user cancels
except Exception as e:
    sys.exit(0)  # Exit gracefully for other unexpected errors

Fhangers = [doc.GetElement(elId) for elId in pipesel]

family_pathCC = path + NewFilename

tg = TransactionGroup(doc, "Place CABLE Brace Family")
tg.Start()

t = Transaction(doc, 'Load CABLE Brace Family')
# Start Transaction
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    target_family = doc.LoadFamily(family_pathCC, fload_handler)

t.Commit()

# Retrieve the specific family symbol by family and type name
familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralStiffener).OfClass(FamilySymbol).ToElements()
target_famtype = None
for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        target_famtype = famtype
        break

if target_famtype:
    t = Transaction(doc, 'Activate and Populate CABLE Brace Family')
    t.Start()
    target_famtype.Activate()
    doc.Regenerate()

    for hanger in Fhangers:
        STName = hanger.GetRodInfo().RodCount
        STName1 = hanger.GetRodInfo()
        if STName == 1:
            ItmDims = hanger.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Rod Length':
                    RodLength = hanger.GetDimensionValue(dta)
                    BraceOffsetZ = RodLength
                elif dta.Name == 'RodLength':  # Check for 'RodLength' if 'Rod Length' isn't found
                    RodLength = hanger.GetDimensionValue(dta)
                    BraceOffsetZ = RodLength
            bounding_box = hanger.get_BoundingBox(None)
            if bounding_box is not None:
                middle_bottom_point = XYZ((bounding_box.Min.X + bounding_box.Max.X) / 2,
                                          (bounding_box.Min.Y + bounding_box.Max.Y) / 2,
                                          bounding_box.Min.Z)

                middle_top_point = XYZ((bounding_box.Min.X + bounding_box.Max.X) / 2,
                                      (bounding_box.Min.Y + bounding_box.Max.Y) / 2,
                                      bounding_box.Max.Z)

            rodloc = STName1.GetRodEndPosition(0)
            valuenum = rodloc.Z
 
            new_insertion_point = XYZ(middle_top_point.X, middle_top_point.Y, (middle_top_point.Z + 0.025716145) - BraceOffsetZ)
            # First brace
            new_family_instance = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
            stretch_brace(new_family_instance, valuenum)
            
            # Second brace, rotated 180 degrees
            new_family_instance2 = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
            axis_start = new_insertion_point
            axis_end = XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)
            rotation_axis = Line.CreateBound(axis_start, axis_end)
            ElementTransformUtils.RotateElement(doc, new_family_instance2.Id, rotation_axis, math.pi)  # 180 degrees
            stretch_brace(new_family_instance2, valuenum)

        if STName > 1:
            RackType = get_parameter_value_by_name_AsValueString(hanger, 'Family')
            bounding_box = hanger.get_BoundingBox(None)
            if bounding_box is not None:
                middle_bottom_point = XYZ((bounding_box.Min.X + bounding_box.Max.X) / 2,
                                          (bounding_box.Min.Y + bounding_box.Max.Y) / 2,
                                          bounding_box.Min.Z)

            for n in range(STName):
                rodloc = STName1.GetRodEndPosition(n)
                valuenum = rodloc.Z
                if RackType == '1.625 Single Strut Trapeze':
                    combined_xyz = XYZ(rodloc.X, rodloc.Y, (middle_bottom_point.Z + 0.25524))
                elif RackType == '1.625 Double Strut Trapeze':
                    combined_xyz = XYZ(rodloc.X, rodloc.Y, (middle_bottom_point.Z + 0.390656))
                else:
                    combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z)  # Fallback

                # First brace
                new_family_instance = doc.Create.NewFamilyInstance(combined_xyz, target_famtype, DB.Structure.StructuralType.NonStructural)
                stretch_brace(new_family_instance, valuenum)
                
                # Second brace, rotated 180 degrees
                new_family_instance2 = doc.Create.NewFamilyInstance(combined_xyz, target_famtype, DB.Structure.StructuralType.NonStructural)
                axis_start = combined_xyz
                axis_end = XYZ(combined_xyz.X, combined_xyz.Y, combined_xyz.Z + 1)
                rotation_axis = Line.CreateBound(axis_start, axis_end)
                ElementTransformUtils.RotateElement(doc, new_family_instance2.Id, rotation_axis, math.pi)  # 180 degrees
                stretch_brace(new_family_instance2, valuenum)

    t.Commit()
tg.Assimilate()