from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, BuiltInCategory, FamilySymbol, LocationCurve, Transaction
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger
import re
from math import atan2, degrees
from fractions import Fraction
import os

path, filename = os.path.split(__file__)
NewFilename = '\Round Wall Sleeve.rfa'

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Wall-Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('6')

with open(filepath, 'r') as f:
    SleeveLength = float(f.read())

# Get the associated level of the active view
level = active_view.GenLevel

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
FamilyName = 'Round Wall Sleeve'
FamilyType = 'Round Wall Sleeve'
# Check if the family is in the project
Fam_is_in_project = any(f.Name == FamilyName for f in families)

family_pathCC = path + NewFilename

t = Transaction(doc, 'Load Wall Sleeve Family')
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

symbName = 'Round Wall Sleeve'

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
        diameter = float(re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)[^\d.]', frac2string, get_parameter_value_by_name_AsString(pipe, 'Overall Size'))) / 12
    else:
        diameter = float(re.sub(r'[^\d.]', '', get_parameter_value_by_name_AsString(pipe, 'Overall Size'))) / 12

    if 0.0 < (diameter *12) < 1.0:
        set_parameter_by_name(new_family_instance, 'Diameter', (2.0 / 12))
    if 1.0 < (diameter *12) < 1.25:
        set_parameter_by_name(new_family_instance, 'Diameter', (2.5 / 12))
    if 1.25 < (diameter *12) < 1.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (3.0 / 12))
    if 1.5 < (diameter *12) < 2.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (4.0 / 12))
    if 2.5 < (diameter *12) < 3.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (5.0 / 12))
    if 3.5 < (diameter *12) < 4.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (6.0 / 12))
    if 4.5 < (diameter *12) < 7.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (8.0 / 12))
    if 7.5 < (diameter *12) < 8.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (10.0 / 12))
    if 8.5 < (diameter *12) < 10.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (12.0 / 12))
    if 10.5 < (diameter *12) < 14.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (16.0 / 12))
    if 14.5 < (diameter *12) < 16.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (18.0 / 12))
    if 16.5 < (diameter *12) < 18.5:
        set_parameter_by_name(new_family_instance, 'Diameter', (20.0 / 12))
    set_parameter_by_name(new_family_instance, 'Length', SleeveLength)
    
    # Get connector locations
    pipe_connectors = list(pipe.ConnectorManager.Connectors)
    connector1, connector2 = pipe_connectors[0], pipe_connectors[1]

    # Calculate distances to the picked_point
    distance1 = picked_point.DistanceTo(connector1.Origin)
    distance2 = picked_point.DistanceTo(connector2.Origin)

    # Determine the nearest connector
    if distance1 < distance2:
        connector1, connector2 = connector1, connector2
    else:
        connector1, connector2 = connector2, connector1

    # Calculate vector components and angle
    vec_x = connector2.Origin.X - connector1.Origin.X
    vec_y = connector2.Origin.Y - connector1.Origin.Y
    angle = atan2(vec_y, vec_x)
    axis = DB.Line.CreateBound(insertion_point, DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1))
    
    # Set rotation on new family placed in model
    DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)
    
    # Set FP parameters on new family placed in model
    set_parameter_by_name(new_family_instance, 'FP_Product Entry', get_parameter_value_by_name_AsString(pipe, 'Overall Size'))
    set_parameter_by_name(new_family_instance, 'FP_Service Name', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Name'))
    set_parameter_by_name(new_family_instance, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Abbreviation'))
    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    schedule_level_param.Set(level.Id)


while True:
    try:
        t = Transaction(doc, 'Place Wall Sleeve Family')
        t.Start()
        
        pipe = select_fabrication_pipe()
        place_and_modify_family(pipe, famsymb)
        
        t.Commit()
        
    except Exception as e:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        break
