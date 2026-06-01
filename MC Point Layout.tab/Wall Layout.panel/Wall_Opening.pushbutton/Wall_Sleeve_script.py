# -*- coding: utf-8 -*-
import clr
import os
import math
from math import atan2

clr.AddReference('System')
import System

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, Family, FamilySymbol, BuiltInCategory,
    LocationCurve, Transaction
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from Parameters.Get_Set_Params import (
    get_parameter_value_by_name_AsString
)

# --------------------------------------------------
# Basic environment
# --------------------------------------------------
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
uiapp = __revit__

script_dir, script_file = os.path.split(__file__)
family_name = 'WS'
family_type = 'WS'
family_path = os.path.join(script_dir, 'WS.rfa')


# --------------------------------------------------
# Family load options
# --------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


# --------------------------------------------------
# Family manager
# --------------------------------------------------
class WSFamilyManager(object):
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

        t = Transaction(self.doc, "Load {} Family".format(self.family_name))
        try:
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
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            TaskDialog.Show("Error", "Error loading family:\n{}".format(str(e)))
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

        t = Transaction(self.doc, "Activate {} Symbol".format(self.family_name))
        try:
            t.Start()
            symbol.Activate()
            self.doc.Regenerate()
            t.Commit()
            return True
        except Exception as e:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            TaskDialog.Show("Error", "Error activating symbol:\n{}".format(str(e)))
            return False

    def get_ready_symbol(self, type_name):
        fam = self.load_family_if_missing()
        if not fam:
            return None

        self.build_symbol_cache()
        sym = self.get_symbol_by_type_name(type_name)
        if not sym:
            TaskDialog.Show("Error", "Type '{}' not found in family '{}'.".format(type_name, self.family_name))
            return None

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


# --------------------------------------------------
# Helpers
# --------------------------------------------------
def set_parameter_by_name(element, parameter_name, value):
    p = element.LookupParameter(parameter_name)
    if p and not p.IsReadOnly:
        p.Set(value)

def get_pipe_centerline(pipe):
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    raise Exception("Selected element does not have a valid centerline.")

def project_point_on_curve(point, curve):
    result = curve.Project(point)
    if not result:
        raise Exception("Could not project picked point onto pipe centerline.")
    return result.XYZPoint

def round_up_to_nearest_quarter(value_feet):
    value_inches = value_feet * 12.0
    rounded_inches = math.ceil(value_inches * 4.0) / 4.0
    return rounded_inches / 12.0

def select_fabrication_pipe():
    pipe_ref = uidoc.Selection.PickObject(ObjectType.Element, "Select an MEP Fabrication Pipe")
    return doc.GetElement(pipe_ref.ElementId)

def pick_point():
    return uidoc.Selection.PickPoint("Pick a point along the centerline of the pipe")


# --------------------------------------------------
# Placement
# --------------------------------------------------
def place_ws_family(pipe, symbol):
    centerline_curve = get_pipe_centerline(pipe)
    picked_point = pick_point()
    projected_point = project_point_on_curve(picked_point, centerline_curve)

    level = doc.GetElement(pipe.LevelId)

    # Better than mixing XY/Z: place directly on pipe centerline
    insertion_point = projected_point

    new_inst = doc.Create.NewFamilyInstance(
        insertion_point,
        symbol,
        DB.Structure.StructuralType.NonStructural
    )

    od_param = pipe.LookupParameter('Outside Diameter')
    if od_param is None:
        raise Exception("Outside Diameter parameter not found.")
    outside_dia = od_param.AsDouble()

    ins_thick = 0.0
    ins_spec_param = pipe.LookupParameter('Insulation Specification')
    if ins_spec_param and ins_spec_param.AsInteger() != 0:
        it_param = pipe.LookupParameter('Insulation Thickness')
        if it_param:
            ins_thick = it_param.AsDouble()

    overall_dia_feet = outside_dia + (2 * ins_thick)
    sleeve_dia = overall_dia_feet + 0.0833333  # +1"
    set_parameter_by_name(new_inst, 'Diameter', round_up_to_nearest_quarter(sleeve_dia))

    pipe_connectors = list(pipe.ConnectorManager.Connectors)
    if len(pipe_connectors) < 2:
        raise Exception("Pipe does not have 2 connectors.")

    c1 = pipe_connectors[0]
    c2 = pipe_connectors[1]

    if picked_point.DistanceTo(c2.Origin) < picked_point.DistanceTo(c1.Origin):
        c1, c2 = c2, c1

    vec_x = c2.Origin.X - c1.Origin.X
    vec_y = c2.Origin.Y - c1.Origin.Y
    angle = atan2(vec_y, vec_x)

    axis = DB.Line.CreateBound(
        insertion_point,
        DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1)
    )
    DB.ElementTransformUtils.RotateElement(doc, new_inst.Id, axis, angle)

    set_parameter_by_name(new_inst, 'FP_Product Entry', str(overall_dia_feet * 12.0))
    set_parameter_by_name(new_inst, 'FP_Service Name',
        get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Name'))
    set_parameter_by_name(new_inst, 'FP_Service Abbreviation',
        get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Abbreviation'))

    sched_level = new_inst.LookupParameter("Schedule Level")
    if sched_level and not sched_level.IsReadOnly:
        sched_level.Set(level.Id)


# --------------------------------------------------
# Main
# --------------------------------------------------
def main():
    family_manager = WSFamilyManager(doc, family_name, family_path)
    ws_symbol = family_manager.get_ready_symbol(family_type)
    if not ws_symbol:
        return

    while True:
        t = Transaction(doc, "Place Trimble Wall Sleeve Family")
        try:
            pipe = select_fabrication_pipe()
            t.Start()
            place_ws_family(pipe, ws_symbol)
            t.Commit()

        except OperationCanceledException:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            break

        except Exception as e:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            TaskDialog.Show("Error", str(e))
            break


if __name__ == '__main__':
    main()