from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, BuiltInCategory, FamilySymbol, LocationCurve, Transaction
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger
import re
from math import atan2, degrees
from fractions import Fraction

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'Pipe Label'
FamilyType = 'Pipe Label'
# Check if the family is in the project
Fam_is_in_project = any(f.Name == FamilyName for f in families)

family_pathCC = r'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES\Pipe Accessories\Sleeves\Pipe Label.rfa'

t = Transaction(doc, 'Load Pipe Label Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC, fload_handler)

collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)

# Filter family symbols by family name and type name
famsymb = None
for fs in collector:
    if fs.Family.Name == FamilyName and fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType:
        famsymb = fs
        break

if famsymb:
    famsymb.Activate()
    doc.Regenerate()

t.Commit()

symbName = 'WS'

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def get_parameter_value_by_name_AsDouble(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def select_fabrication_pipe():
    selection = uidoc.Selection
    pipe_ref = selection.PickObject(ObjectType.Element, "Select an MEP Fabrication Pipe")
    pipe = doc.GetElement(pipe_ref.ElementId)
    return pipe

def pick_point():
    picked_point = uidoc.Selection.PickPoint("Pick a point along the centerline of the pipe")
    return picked_point

def get_pipe_centerline(pipe):
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    else:
        raise Exception("The selected element does not have a valid centerline.")

def project_point_on_curve(point, curve):
    result = curve.Project(point)
    return result.XYZPoint

def place_and_modify_family(pipe, famsymb):
    centerline_curve = get_pipe_centerline(pipe)
    picked_point = pick_point()
    
    # Project the picked point onto the pipe centerline to get the Z coordinate
    projected_point = project_point_on_curve(picked_point, centerline_curve)
    
    # Create the insertion point using the X and Y from the picked point and Z from the projected point
    insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)
    
    new_family_instance = doc.Create.NewFamilyInstance(insertion_point, famsymb, DB.Structure.StructuralType.NonStructural)

    def frac2string(s):
        i, f = s.groups(0)
        f = Fraction(f)
        return str(int(i) + float(f))

    if '/' in get_parameter_value_by_name_AsString(pipe, 'Overall Size'):
        diameter = float(re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)[^\d.]', frac2string, get_parameter_value_by_name_AsString(pipe, 'Overall Size'))) / 12 + 0.0833333
    else:
        diameter = float(re.sub(r'[^\d.]', '', get_parameter_value_by_name_AsString(pipe, 'Overall Size'))) / 12 + 0.0833333
    set_parameter_by_name(new_family_instance, "Diameter", diameter)
    
    # Get connector location and angle on pipe
    pipe_connector = pipe.ConnectorManager.Connectors
    connector1, connector2 = list(pipe_connector)
    vec_x = connector2.Origin.X - connector1.Origin.X
    vec_y = connector2.Origin.Y - connector1.Origin.Y
    angle = atan2(vec_y, vec_x)
    axis = DB.Line.CreateBound(insertion_point, DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1))
    
    # Set rotation on new family placed in model
    DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)
    
    # Set FP parameters on new family placed in model
    set_parameter_by_name(new_family_instance, 'FP_Service Name', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Name'))
    set_parameter_by_name(new_family_instance, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Abbreviation'))

while True:
    try:
        t = Transaction(doc, 'Place Pipe Label Family')
        t.Start()
        
        pipe = select_fabrication_pipe()
        place_and_modify_family(pipe, famsymb)
        
        t.Commit()
        
    except Exception as e:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        break
