from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, LocationCurve, Transaction
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString
from Parameters.Add_SharedParameters import Shared_Params
from math import atan2
from fractions import Fraction
import clr
import sys
import os
import re

Shared_Params()

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

FamilyName = 'RDS'
FamilyType = 'RDS'


# --------------------------------------------------
# Family load options / manager
# --------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


class RDSFamilyManager(object):
    def __init__(self, document, target_family_name, target_family_path):
        self.doc = document
        self.family_name = target_family_name
        self.family_path = target_family_path
        self.family = None
        self.symbol_cache = {}

    def get_family_by_name(self):
        if self.family and self.family.IsValidObject:
            return self.family

        for fam in FilteredElementCollector(self.doc).OfClass(Family):
            if fam.Name == self.family_name:
                self.family = fam
                return fam

        self.family = None
        return None

    def load_family_if_missing(self):
        fam = self.get_family_by_name()
        if fam:
            return fam

        if not os.path.exists(self.family_path):
            TaskDialog.Show("Error", "Family file not found:\n{}".format(self.family_path))
            return None

        t = None
        try:
            t = Transaction(self.doc, "Load {} Family".format(self.family_name))
            t.Start()

            loaded_family_ref = clr.Reference[Family]()
            result = self.doc.LoadFamily(
                self.family_path,
                FamilyLoaderOptionsHandler(),
                loaded_family_ref
            )

            t.Commit()

            if result and loaded_family_ref.Value:
                self.family = loaded_family_ref.Value
                return self.family

            return self.get_family_by_name()

        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            TaskDialog.Show("Error", "Family load error:\n{}".format(str(e)))
            return None
        finally:
            if t:
                t.Dispose()

    def get_symbol_name(self, symbol):
        try:
            if symbol.Name:
                return symbol.Name
        except:
            pass

        try:
            p = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if p:
                return p.AsString()
        except:
            pass

        return None

    def build_symbol_cache(self):
        self.symbol_cache = {}

        fam = self.get_family_by_name()
        if not fam:
            return

        for symbol_id in fam.GetFamilySymbolIds():
            sym = self.doc.GetElement(symbol_id)
            if sym:
                type_name = self.get_symbol_name(sym)
                if type_name:
                    self.symbol_cache[type_name.strip().upper()] = sym

    def get_symbol_by_type_name(self, type_name):
        if not type_name:
            return None

        if not self.symbol_cache:
            self.build_symbol_cache()

        return self.symbol_cache.get(type_name.strip().upper())

    def activate_symbol_if_needed(self, symbol):
        if not symbol:
            return False

        if symbol.IsActive:
            return True

        t = None
        try:
            t = Transaction(self.doc, "Activate {} Symbol".format(self.family_name))
            t.Start()
            symbol.Activate()
            self.doc.Regenerate()
            t.Commit()
            return True
        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            TaskDialog.Show("Error", "Family symbol activation error:\n{}".format(str(e)))
            return False
        finally:
            if t:
                t.Dispose()

    def get_ready_symbol(self, type_name):
        fam = self.load_family_if_missing()
        if not fam:
            return None

        self.build_symbol_cache()

        sym = self.get_symbol_by_type_name(type_name)
        if not sym:
            TaskDialog.Show(
                "Error",
                "Type '{}' not found in family '{}'.".format(type_name, self.family_name)
            )
            return None

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


# --------------------------------------------------
# Settings / family path
# --------------------------------------------------
path, filename = os.path.split(__file__)
family_filename = 'RDS.rfa'
family_pathCC = os.path.join(path, family_filename)

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Duct-Wall-Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

try:
    with open(filepath, 'r') as f:
        AnnularSpace = float(f.read())
except:
    AnnularSpace = 1.0
    TaskDialog.Show("Warning", "Failed to read annular space. Using default 1 inch.")


family_manager = RDSFamilyManager(doc, FamilyName, family_pathCC)
famsymb = family_manager.get_ready_symbol(FamilyType)

if not famsymb:
    sys.exit(1)


# --------------------------------------------------
# Helpers
# --------------------------------------------------
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

def get_metal_diameter_from_duct(duct, connector=None):
    try:
        if connector and connector.Shape == DB.ConnectorProfileType.Round:
            return connector.Radius * 2.0
    except:
        pass

    overall_size = get_parameter_value_by_name_AsString(duct, 'Overall Size')
    if not overall_size:
        raise Exception("Could not read duct size.")

    def frac2string(s):
        i, f = s.groups(0)
        f = Fraction(f)
        return str(int(i) + float(f))

    if '/' in overall_size:
        return float(
            re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)[^\d.]', frac2string, overall_size)
        ) / 12.0
    else:
        return float(re.sub(r'[^\d.]', '', overall_size)) / 12.0


# --------------------------------------------------
# Placement
# --------------------------------------------------
def place_and_modify_family(duct, famsymb):
    try:
        centerline_curve = get_pipe_centerline(duct)
        picked_point = pick_point()
        if not picked_point:
            return False

        projected_point = project_point_on_curve(picked_point, centerline_curve)
        insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)

        level = doc.GetElement(duct.LevelId)

        new_family_instance = doc.Create.NewFamilyInstance(
            insertion_point,
            famsymb,
            DB.Structure.StructuralType.NonStructural
        )

        pipe_connectors = list(duct.ConnectorManager.Connectors)
        if len(pipe_connectors) < 2:
            raise Exception("Selected duct does not have at least 2 connectors.")

        connector1, connector2 = pipe_connectors[0], pipe_connectors[1]

        distance1 = picked_point.DistanceTo(connector1.Origin)
        distance2 = picked_point.DistanceTo(connector2.Origin)
        if distance2 < distance1:
            connector1, connector2 = connector2, connector1

        diameter = get_metal_diameter_from_duct(duct, connector1) + (AnnularSpace / 12.0)
        set_parameter_by_name(new_family_instance, 'Diameter', diameter)

        vec_x = connector2.Origin.X - connector1.Origin.X
        vec_y = connector2.Origin.Y - connector1.Origin.Y
        angle = atan2(vec_y, vec_x)

        axis = DB.Line.CreateBound(
            insertion_point,
            DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1)
        )

        DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)

        set_parameter_by_name(
            new_family_instance,
            'FP_Service Name',
            get_parameter_value_by_name_AsString(duct, 'Fabrication Service Name')
        )

        schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
        if schedule_level_param and not schedule_level_param.IsReadOnly:
            schedule_level_param.Set(level.Id)

        return True

    except Exception as e:
        raise Exception("Family placement error: {}".format(e))


# --------------------------------------------------
# Main
# --------------------------------------------------
while True:
    t = None
    try:
        duct = select_fabrication_duct()
        if not duct:
            break

        if duct.ItemCustomId == 1:
            TaskDialog.Show("Warning", "You selected a rectangular duct.")
            break

        t = Transaction(doc, 'Place Trimble Wall Sleeve Family')
        t.Start()

        if not place_and_modify_family(duct, famsymb):
            t.RollBack()
            break

        t.Commit()

    except OperationCanceledException:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        break

    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        TaskDialog.Show("Error", "Error during operation:\n{}".format(e))
        break

    finally:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if t:
            t.Dispose()