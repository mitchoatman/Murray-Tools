import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, Structure, XYZ, FabricationPart, FabricationConfiguration, TransactionGroup, BuiltInParameter, Line, ElementTransformUtils
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString, get_parameter_value_by_name_AsDouble, set_parameter_by_name, get_parameter_value_by_name_AsString
from Autodesk.Revit.UI import TaskDialog
import math
import os
import sys

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

def stretch_brace():
    try:
        # This section extends the brace fam to top of hanger rod elevation after its placed.
        # Writes data to TOS Parameter
        set_parameter_by_name(new_family_instance, "Top of Steel", valuenum)
        # Reads brace angle
        BraceAngle = get_parameter_value_by_name_AsDouble(new_family_instance, "BraceMainAngle")
        sinofangle = math.sin(BraceAngle)
        # Reads brace elevation
        BraceElevation = get_parameter_value_by_name_AsDouble(new_family_instance, 'Offset from Host')
        # Equation to get the new hypotenus
        Height = ((valuenum - BraceElevation) - 0.2330)
        newhypotenus = ((Height / sinofangle) - 0.175)
        if newhypotenus < 0:
            newhypotenus = 1
        # Writes new Brace length to parameter
        set_parameter_by_name(new_family_instance, "BraceLength", newhypotenus)
        # Writes Level into Brace
        set_parameter_by_name(new_family_instance, "ISAT Brace Level", HangerLevel)
        # Only set FP_Service Name if parameter exists
        if new_family_instance.LookupParameter("FP_Service Name"):
            set_parameter_by_name(new_family_instance, "FP_Service Name", HangerService)
    except:
        pass

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

        RackType = get_parameter_value_by_name_AsValueString(hanger, 'Family')
        rack_type_lower = RackType.lower()

        if not 'strap' in rack_type_lower:
            bounding_box = hanger.get_BoundingBox(None)
            if bounding_box is not None:
                middle_bottom_point = XYZ((bounding_box.Min.X + bounding_box.Max.X) / 2,
                                          (bounding_box.Min.Y + bounding_box.Max.Y) / 2,
                                          bounding_box.Min.Z)

            if bounding_box is not None:
                middle_top_point = XYZ((bounding_box.Min.X + bounding_box.Max.X) / 2,
                                          (bounding_box.Min.Y + bounding_box.Max.Y) / 2,
                                          bounding_box.Max.Z)

            HangerLevel = get_parameter_value_by_name_AsValueString(hanger, 'Reference Level')
            HangerService = get_parameter_value_by_name_AsString(hanger, 'Fabrication Service Name')

            STName = hanger.GetRodInfo().RodCount
            STName1 = hanger.GetRodInfo()
            

            ItmDims = hanger.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Rod Length':
                    RodLength = hanger.GetDimensionValue(dta)
                    BraceOffsetZ = RodLength
                elif dta.Name == 'RodLength':  # Check for 'RodLength' if 'Rod Length' isn't found
                    RodLength = hanger.GetDimensionValue(dta)
                    BraceOffsetZ = RodLength
                elif dta.Name == 'Rod Extn Above':
                    RodLength = hanger.GetDimensionValue(dta)
                    BraceOffsetZ = RodLength
                # if dta.Name == 'Total Height':
                    # HangerHeight = hanger.GetDimensionValue(dta)
                    # BraceOffsetZ = HangerHeight + 0.01041666
                # if dta.Name == 'Weld Lug Height':
                    # HangerHeight = hanger.GetDimensionValue(dta)
                    # BraceOffsetZ = HangerHeight + 0.1197916  

            if STName == 1:
                ItmDims = hanger.GetDimensions()
                for dta in ItmDims:
                    if dta.Name == 'Rod Length':
                        RodLength = hanger.GetDimensionValue(dta)
                        BraceOffsetZ = RodLength
                    elif dta.Name == 'RodLength':  # Check for 'RodLength' if 'Rod Length' isn't found
                        RodLength = hanger.GetDimensionValue(dta)
                        BraceOffsetZ = RodLength

                rodloc = STName1.GetRodEndPosition(0)
                valuenum = rodloc.Z

                
     
                new_insertion_point = XYZ(middle_top_point.X, middle_top_point.Y, (middle_top_point.Z + 0.025716145) - BraceOffsetZ)
                # First brace
                new_family_instance = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
                stretch_brace()
                
                # Second brace, rotated 180 degrees
                new_family_instance = doc.Create.NewFamilyInstance(new_insertion_point, target_famtype, DB.Structure.StructuralType.NonStructural)
                axis_start = new_insertion_point
                axis_end = XYZ(new_insertion_point.X, new_insertion_point.Y, new_insertion_point.Z + 1)
                rotation_axis = Line.CreateBound(axis_start, axis_end)
                ElementTransformUtils.RotateElement(doc, new_family_instance.Id, rotation_axis, math.pi)  # 180 degrees
                stretch_brace()

            if STName > 1:
                error_displayed = False

                for n in range(STName):
                    rodloc = STName1.GetRodEndPosition(n)
                    valuenum = rodloc.Z
                    combined_xyz = XYZ(rodloc.X, rodloc.Y, rodloc.Z)

                    if "1.625" in rack_type_lower:
                        if "single" in rack_type_lower:
                            combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.25524)
                            combined_xyz2 = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.28649)
                        elif "double" in rack_type_lower:
                            combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.390656)
                            combined_xyz2 = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.421906)
                    elif "unistrut" in rack_type_lower:
                        if "038" in rack_type_lower or "1-5/8" in rack_type_lower:
                            combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.25)
                            combined_xyz2 = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.28125)
                        elif "050" in rack_type_lower:
                            if "double" in rack_type_lower or "doublestrut" in rack_type_lower:
                                combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.4165)
                                combined_xyz2 = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.44775)
                            else:
                                combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.2815)
                                combined_xyz2 = XYZ(rodloc.X, rodloc.Y, middle_bottom_point.Z + 0.31275)

                    elif "seismic" in rack_type_lower:
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_top_point.Z - BraceOffsetZ + 0.1615)
                        combined_xyz2 = XYZ(rodloc.X, rodloc.Y, middle_top_point.Z - BraceOffsetZ + 0.19275)

                    # First brace
                    new_family_instance = doc.Create.NewFamilyInstance(combined_xyz, target_famtype, DB.Structure.StructuralType.NonStructural)
                    stretch_brace()
                    
                    # Second brace, rotated 180 degrees
                    new_family_instance = doc.Create.NewFamilyInstance(combined_xyz2, target_famtype, DB.Structure.StructuralType.NonStructural)
                    axis_start = combined_xyz
                    axis_end = XYZ(combined_xyz.X, combined_xyz.Y, combined_xyz.Z + 1)
                    rotation_axis = Line.CreateBound(axis_start, axis_end)
                    ElementTransformUtils.RotateElement(doc, new_family_instance.Id, rotation_axis, math.pi)  # 180 degrees
                    stretch_brace()
        else:
            TaskDialog.Show("Error", "Cannot Place Seismic on that Support!")

    t.Commit()
tg.Assimilate()