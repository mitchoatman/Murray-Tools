
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, BuiltInCategory, FamilySymbol, LocationCurve,ElementCategoryFilter, BuiltInCategory, ElementClassFilter, BuiltInParameter, \
                                ElementId, ElementParameterFilter, ParameterValueProvider, FilterStringRule, FilterStringEquals, LogicalAndFilter, Transaction, FamilyInstance
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, \
                                        get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsValueString, get_parameter_value_by_name_AsDouble
from Autodesk.Revit.UI.Selection import ObjectType
import re
from math import atan2, degrees
from fractions import Fraction
from SharedParam.Add_Parameters import Shared_Params
import os



Shared_Params()

path, filename = os.path.split(__file__)
NewFilename = '\Pipe Label.rfa'

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

# Get the associated level of the active view
level = curview.GenLevel

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

family_pathCC = path + NewFilename

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

symbName = 'Pipe Label'

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
    set_parameter_by_name(new_family_instance, "Diameter", diameter)
    
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
    set_parameter_by_name(new_family_instance, 'FP_Service Name', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Name'))
    set_parameter_by_name(new_family_instance, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Abbreviation'))
    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    schedule_level_param.Set(level.Id)

def set_pipe_label_size():

    # Create a filter for pipe accessories
    pipe_accessory_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory)

    # Create a filter for family instances (optional, as pipe accessories are already family instances)
    family_instance_filter = ElementClassFilter(FamilyInstance)

    # Create a parameter value provider for the family name
    provider = ParameterValueProvider(ElementId(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM))

    # Define the rule to filter by the family name "Pipe Label"
    # Remove False for new Revit versions
    if RevitINT > 2022:
        rule = FilterStringRule(provider, FilterStringEquals(), "Pipe Label")
    else:
        rule = FilterStringRule(provider, FilterStringEquals(), "Pipe Label", False)

    # Create a parameter filter using the rule
    family_name_filter = ElementParameterFilter(rule)

    # Combine the filters
    filter = LogicalAndFilter(pipe_accessory_filter, family_name_filter)

    # Collect all elements that match the filter
    pipe_labels = FilteredElementCollector(doc).WherePasses(filter).ToElements()

    t = Transaction(doc, "Set Pipe Label type")
    t.Start()
    # Iterate over elements and fetch parameter values
    for pipe_label in pipe_labels:
        try:
            diameter = (get_parameter_value_by_name_AsDouble(pipe_label, 'Diameter') * 12)
            # print diameter
            # print diameter <= 0.5
            if diameter <= 0.50:
                set_parameter_by_name(pipe_label, 'FP_Product Entry', 'AA')
            if 0.50 < diameter < 1.001:
                set_parameter_by_name(pipe_label, 'FP_Product Entry', 'A')
            if 1.00 < diameter < 2.376:
                set_parameter_by_name(pipe_label, 'FP_Product Entry', 'B')
            if 2.375 < diameter < 3.251:
                set_parameter_by_name(pipe_label, 'FP_Product Entry', 'C')
            if 3.25 < diameter < 4.501:
                set_parameter_by_name(pipe_label, 'FP_Product Entry', 'D')
            if 4.50 < diameter < 5.876:
                set_parameter_by_name(pipe_label, 'FP_Product Entry', 'E')
            if 5.875 < diameter < 7.876:
                set_parameter_by_name(pipe_label, 'FP_Product Entry', 'F')
            if 7.875 < diameter < 9.876:
                set_parameter_by_name(pipe_label, 'FP_Product Entry', 'G')
            if diameter > 9.876:
                set_parameter_by_name(pipe_label, 'FP_Product Entry', 'H')
        except:
            pass
    t.Commit()


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

set_pipe_label_size()