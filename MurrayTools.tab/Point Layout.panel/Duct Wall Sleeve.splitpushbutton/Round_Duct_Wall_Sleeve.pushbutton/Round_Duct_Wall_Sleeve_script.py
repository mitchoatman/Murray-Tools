from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, BuiltInCategory, FamilySymbol, LocationCurve, Transaction, ViewType
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsDouble
from Autodesk.Revit.UI import TaskDialog
from math import atan2
from fractions import Fraction
from Parameters.Add_SharedParameters import Shared_Params
import clr, sys, os, re

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
        FamilyName = 'RDS'
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
NewFilename = '\RDS.rfa'
family_pathCC = path + NewFilename
folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Duct-Wall-Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

with open(filepath, 'r') as f:
    AnnularSpace = float(f.read())

try:
    load_family(family_pathCC)
except Exception as e:
    print("Family load failed: {}".format(e))
    sys.exit(1)

collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).OfClass(FamilySymbol)
FamilyName = 'RDS'
FamilyType = 'RDS'
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

def select_fabrication_duct():
    try:
        pipe_ref = uidoc.Selection.PickObject(ObjectType.Element, "Select a round MEP Fabrication Duct")
        return doc.GetElement(pipe_ref.ElementId)
    except:
        return None

def pick_point():
    try:
        return uidoc.Selection.PickPoint("Pick a point along the centerline of the Duct")
    except:
        return None

def get_pipe_centerline(duct):
    pipe_location = duct.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    raise Exception("Invalid duct centerline")

def project_point_on_curve(point, curve):
    result = curve.Project(point)
    return result.XYZPoint

def place_and_modify_family(duct, famsymb):
    try:
        centerline_curve = get_pipe_centerline(duct)
        picked_point = pick_point()
        if not picked_point:
            return False  # Quietly exit if user cancels point selection
        projected_point = project_point_on_curve(picked_point, centerline_curve)
        insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)

        # Get reference level from duct
        level = doc.GetElement(duct.LevelId)
  
        new_family_instance = doc.Create.NewFamilyInstance(insertion_point, famsymb, DB.Structure.StructuralType.NonStructural)

        def frac2string(s):
            i, f = s.groups(0)
            f = Fraction(f)
            return str(int(i) + float(f))

        if '/' in get_parameter_value_by_name_AsString(duct, 'Overall Size'):
            diameter = float(re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)[^\d.]', frac2string, get_parameter_value_by_name_AsString(duct, 'Overall Size'))) /12 + (AnnularSpace / 12)
        else:
            diameter = float(re.sub(r'[^\d.]', '', get_parameter_value_by_name_AsString(duct, 'Overall Size'))) /12 + (AnnularSpace / 12)
        set_parameter_by_name(new_family_instance, 'Diameter', diameter)

        pipe_connectors = list(duct.ConnectorManager.Connectors)
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
        set_parameter_by_name(new_family_instance, 'FP_Service Name', get_parameter_value_by_name_AsString(duct, 'Fabrication Service Name'))
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
        duct = select_fabrication_duct()
        if not duct:
            break  # Quietly exit if user cancels duct selection
        if not place_and_modify_family(duct, famsymb):
            break  # Quietly exit if user cancels point selection
        if duct.ItemCustomId == 1:
            print 'You selected a rect duct fool!'
            break  # Quietly exit if user selects rect duct
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