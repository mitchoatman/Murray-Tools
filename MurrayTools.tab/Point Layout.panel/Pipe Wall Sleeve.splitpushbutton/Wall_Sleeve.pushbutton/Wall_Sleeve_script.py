from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, BuiltInCategory, FamilySymbol, LocationCurve, Transaction, ViewType
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger
from Autodesk.Revit.UI import TaskDialog
import re, math, os, clr, sys
from math import atan2, degrees
from fractions import Fraction
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

class FamilyLoadOptions(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues[0] = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues[0] = False
        return True

def load_family(family_path):
    t = None
    try:
        t = Transaction(doc, 'Load Trimble Wall Sleeve Family')
        t.Start()
        families = FilteredElementCollector(doc).OfClass(Family)
        FamilyName = 'WS'
        if not any(f.Name == FamilyName for f in families):
            load_options = FamilyLoadOptions()
            loaded_family = clr.StrongBox[DB.Family]()
            success = doc.LoadFamily(family_path, load_options, loaded_family)
            if not success or not loaded_family.Value:
                raise Exception("Failed to load family")
        t.Commit()
    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise Exception("Family load error: {}".format(e))
    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()

path, filename = os.path.split(__file__)
NewFilename = '\WS.rfa'
family_pathCC = path + NewFilename

try:
    load_family(family_pathCC)
except Exception as e:
    print("Family load failed: {}".format(e))
    sys.exit(1)

collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
FamilyName = 'WS'
FamilyType = 'WS'
famsymb = None
for fs in collector:
    if fs.Family.Name == FamilyName and fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType:
        famsymb = fs
        break

if famsymb:
    t = None
    try:
        t = Transaction(doc, 'Activate Family Symbol')
        t.Start()
        famsymb.Activate()
        doc.Regenerate()
        t.Commit()
    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        print("Family symbol activation error: {}".format(e))
        sys.exit(1)
    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def get_parameter_value_by_name_AsDouble(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def select_fabrication_pipe():
    try:
        pipe_ref = uidoc.Selection.PickObject(ObjectType.Element, "Select an MEP Fabrication Pipe")
        return doc.GetElement(pipe_ref.ElementId)
    except:
        return None

def pick_point():
    try:
        return uidoc.Selection.PickPoint("Pick a point along the centerline of the pipe")
    except:
        return None

def get_pipe_centerline(pipe):
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    raise Exception("Invalid pipe centerline")

def project_point_on_curve(point, curve):
    result = curve.Project(point)
    return result.XYZPoint

def inches_to_frac_string(inches):
    if inches < 0:
        raise ValueError("Negative inches not supported")
    whole = int(inches)
    frac_part = inches - whole
    if abs(frac_part) < 1e-6:
        return str(whole) + '"'
    frac = Fraction(frac_part).limit_denominator(8)
    if frac.denominator == 1:
        return str(whole + frac.numerator) + '"'
    else:
        return str(whole) + " " + str(frac.numerator) + "/" + str(frac.denominator) + '"'

def round_up_to_nearest_quarter(value):
    value_in_inches = value * 12
    rounded_value_in_inches = math.ceil(value_in_inches * 4) / 4
    return rounded_value_in_inches / 12

def place_and_modify_family(pipe, famsymb):
    try:
        centerline_curve = get_pipe_centerline(pipe)
        picked_point = pick_point()
        if not picked_point:
            return False  # Quietly exit if user cancels point selection
        projected_point = project_point_on_curve(picked_point, centerline_curve)
        insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)
        
        # Get reference level from pipe
        level = doc.GetElement(pipe.LevelId)
        
        new_family_instance = doc.Create.NewFamilyInstance(insertion_point, famsymb, level, DB.Structure.StructuralType.NonStructural)

        # Calculate overall size from Outside Diameter + 2 * Insulation Thickness
        od_param = pipe.LookupParameter('Outside Diameter')
        if od_param is None:
            raise Exception("Outside Diameter parameter not found on the selected pipe.")
        outside_dia = od_param.AsDouble()  # in feet (internal units)

        ins_spec_param = pipe.LookupParameter('Insulation Specification')
        has_insulation = False
        if ins_spec_param is not None:
            has_insulation = ins_spec_param.AsInteger() != 0

        ins_thick = 0.0
        if has_insulation:
            it_param = pipe.LookupParameter('Insulation Thickness')
            if it_param is not None:
                ins_thick = it_param.AsDouble()  # in feet (internal units)

        overall_dia_feet = outside_dia + 2 * ins_thick
        diameter = overall_dia_feet + 0.0833333
        set_parameter_by_name(new_family_instance, 'Diameter', round_up_to_nearest_quarter(diameter))

        # Format overall size string for FP_Product Entry
        set_parameter_by_name(new_family_instance, 'FP_Product Entry', str(overall_dia_feet * 12) + ' - Pipe OD')

        pipe_connectors = list(pipe.ConnectorManager.Connectors)
        connector1, connector2 = pipe_connectors[0], pipe_connectors[1]
        distance1 = picked_point.DistanceTo(connector1.Origin)
        distance2 = picked_point.DistanceTo(connector2.Origin)
        if distance1 < distance2:
            connector1, connector2 = connector1, connector2
        else:
            connector1, connector2 = connector2, connector1

        vec_x = connector2.Origin.X - connector1.Origin.X
        vec_y = connector2.Origin.Y - connector1.Origin.Y
        angle = atan2(vec_y, vec_x)
        axis = DB.Line.CreateBound(insertion_point, DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1))
        
        DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)
        set_parameter_by_name(new_family_instance, 'FP_Service Name', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Name'))
        set_parameter_by_name(new_family_instance, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Abbreviation'))
        schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
        schedule_level_param.Set(level.Id)
        return True
    except Exception as e:
        raise Exception("Family placement error: {}".format(e))

from Autodesk.Revit.Exceptions import OperationCanceledException

while True:
    t = None
    try:
        t = Transaction(doc, 'Place Trimble Wall Sleeve Family')
        t.Start()
        pipe = select_fabrication_pipe()
        if not pipe:
            break  # Quietly exit if user cancels pipe selection
        if not place_and_modify_family(pipe, famsymb):
            break  # Quietly exit if user cancels point selection
        t.Commit()
    except OperationCanceledException:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        break
    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        print("Error during operation: {}".format(e))
        break
    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()