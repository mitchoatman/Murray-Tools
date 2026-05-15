# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
import System

from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    LocationCurve,
    Transaction
)
from Autodesk.Revit.Exceptions import OperationCanceledException

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsInteger
)
from Parameters.Add_SharedParameters import Shared_Params

import re
import os
from math import atan2
from fractions import Fraction

Shared_Params()

# --------------------------------------------------
# Basic environment
# --------------------------------------------------
path, filename = os.path.split(__file__)
family_pathCC = os.path.join(path, 'Round Wall Sleeve.rfa')

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
active_view = doc.ActiveView

FamilyName = 'Round Wall Sleeve'
FamilyType = 'Round Wall Sleeve'

# --------------------------------------------------
# Helpers
# --------------------------------------------------
def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        pass


# --------------------------------------------------
# Sleeve length storage
# --------------------------------------------------
folder_name = r"c:\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('6')

with open(filepath, 'r') as f:
    SleeveLength = float(f.read())

try:
    level = active_view.GenLevel
except:
    level = None


# --------------------------------------------------
# Family load options
# --------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Project
        overwriteParameterValues.Value = False
        return True


def shared_family_dialog_fallback(sender, args):
    try:
        if isinstance(args, TaskDialogShowingEventArgs):
            msg = (args.Message or "").lower()
            dialog_id = (args.DialogId or "").lower()

            if ("shared" in msg and "already exists" in msg and "project" in msg) \
               or ("shared" in dialog_id and "family" in dialog_id):
                args.OverrideResult(1003)
    except:
        pass


# --------------------------------------------------
# Family manager
# --------------------------------------------------
class RoundWallSleeveFamilyManager(object):
    def __init__(self, document, family_name, family_path):
        self.doc = document
        self.family_name = family_name
        self.family_path = family_path
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

    def load_family_if_missing(self):
        fam = self.get_family_by_name()
        if fam:
            return fam

        if not os.path.exists(self.family_path):
            show_message("Error", "Family file not found:\n{}".format(self.family_path))
            return None

        t = None
        uiapp.DialogBoxShowing += shared_family_dialog_fallback
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
            show_message("Error", "Family load error:\n{}".format(str(e)))
            return None

        finally:
            uiapp.DialogBoxShowing -= shared_family_dialog_fallback
            if t:
                t.Dispose()

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
            show_message("Error", "Family symbol activation error:\n{}".format(str(e)))
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
            show_message(
                "Error",
                "Type '{}' not found in family '{}'.".format(type_name, self.family_name)
            )
            return None

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


family_manager = RoundWallSleeveFamilyManager(doc, FamilyName, family_pathCC)
famsymb = family_manager.get_ready_symbol(FamilyType)

if not famsymb:
    raise Exception("Family symbol '{}' was not found or could not be activated.".format(FamilyType))


# --------------------------------------------------
# Diameter map
# --------------------------------------------------
DIAMETER_MAP = {
    (0.0, 1.0): 2.0, (1.0, 1.25): 2.5, (1.25, 1.5): 3.0,
    (1.5, 2.5): 4.0, (2.5, 3.5): 5.0, (3.5, 4.5): 6.0,
    (4.5, 7.5): 8.0, (7.5, 8.5): 10.0, (8.5, 10.5): 12.0,
    (10.5, 14.5): 16.0, (14.5, 16.5): 18.0, (16.5, 18.5): 20.0,
    (18.5, 20.5): 22.0, (20.5, 22.5): 24.0, (22.5, 24.5): 26.0,
    (24.5, 26.5): 28.0, (26.5, 28.5): 30.0, (28.5, 30.5): 32.0,
    (30.5, 32.5): 34.0, (32.5, 34.5): 36.0
}


# --------------------------------------------------
# Selection / geometry
# --------------------------------------------------
def select_fabrication_pipe():
    return doc.GetElement(
        uidoc.Selection.PickObject(
            ObjectType.Element,
            "Select an MEP Fabrication Pipe"
        ).ElementId
    )


def pick_point():
    return uidoc.Selection.PickPoint("Pick a point along the centerline of the pipe")


def get_pipe_centerline(pipe):
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    raise Exception("The selected element does not have a valid centerline.")


def project_point_on_curve(point, curve):
    result = curve.Project(point)
    if not result:
        raise Exception("Could not project point onto pipe centerline.")
    return result.XYZPoint


def get_diameter_from_size(pipe_diameter):
    pipe_diameter *= 12.0
    for (min_val, max_val), sleeve_size in DIAMETER_MAP.items():
        if min_val < pipe_diameter < max_val:
            return sleeve_size / 12.0
    return 2.0 / 12.0


def create_family_instance(insertion_point, symbol):
    try:
        if level:
            return doc.Create.NewFamilyInstance(
                insertion_point,
                symbol,
                level,
                DB.Structure.StructuralType.NonStructural
            )
    except:
        pass

    return doc.Create.NewFamilyInstance(
        insertion_point,
        symbol,
        DB.Structure.StructuralType.NonStructural
    )


def place_and_modify_family(pipe, famsymb):
    centerline_curve = get_pipe_centerline(pipe)
    picked_point = pick_point()
    projected_point = project_point_on_curve(picked_point, centerline_curve)
    insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)

    new_family_instance = create_family_instance(insertion_point, famsymb)
    if not new_family_instance:
        raise Exception("Failed to create family instance.")

    overall_size = get_parameter_value_by_name_AsString(pipe, 'Overall Size')
    cleaned_size = re.sub(r'["]', '', overall_size.strip())

    try:
        diameter = float(cleaned_size)
    except ValueError:
        match = re.match(r'(?:(\d+)[-\s])?(\d+/\d+)', cleaned_size)
        if match:
            integer_part, fraction_part = match.groups()
            diameter = float(Fraction(fraction_part))
            if integer_part:
                diameter += float(integer_part)
        else:
            show_message(
                "Warning",
                "Could not parse diameter '{}'. Defaulting to 0.5\".".format(overall_size)
            )
            diameter = 0.5

    diameter = diameter / 12.0
    set_parameter_by_name(new_family_instance, 'Diameter', get_diameter_from_size(diameter))
    set_parameter_by_name(new_family_instance, 'Length', SleeveLength)

    connectors = list(pipe.ConnectorManager.Connectors)
    if len(connectors) < 2:
        raise Exception("Selected pipe does not have enough connectors.")

    conn1, conn2 = connectors[0], connectors[1]
    nearest_conn = min([conn1, conn2], key=lambda c: picked_point.DistanceTo(c.Origin))
    other_conn = conn2 if nearest_conn == conn1 else conn1

    vec = other_conn.Origin - nearest_conn.Origin
    angle = atan2(vec.Y, vec.X)
    axis = DB.Line.CreateBound(
        insertion_point,
        DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1.0)
    )
    DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)

    params = {
        'FP_Product Entry': 'Overall Size',
        'FP_Service Name': 'Fabrication Service Name',
        'FP_Service Abbreviation': 'Fabrication Service Abbreviation'
    }

    for fam_param, pipe_param in params.items():
        set_parameter_by_name(
            new_family_instance,
            fam_param,
            get_parameter_value_by_name_AsString(pipe, pipe_param)
        )

    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly and level:
        schedule_level_param.Set(level.Id)


# --------------------------------------------------
# Main execution loop
# --------------------------------------------------
while True:
    t = None
    try:
        pipe = select_fabrication_pipe()

        t = Transaction(doc, 'Place Wall Sleeve Family')
        t.Start()
        place_and_modify_family(pipe, famsymb)
        t.Commit()

    except OperationCanceledException:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        break

    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        show_message("Error", str(e))
        break

    finally:
        if t:
            t.Dispose()