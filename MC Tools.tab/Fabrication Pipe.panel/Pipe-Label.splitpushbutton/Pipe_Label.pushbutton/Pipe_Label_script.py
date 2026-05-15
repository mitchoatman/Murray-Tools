# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
import System

import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    BuiltInCategory,
    FamilySymbol,
    LocationCurve,
    ElementCategoryFilter,
    ElementClassFilter,
    BuiltInParameter,
    ElementId,
    ElementParameterFilter,
    ParameterValueProvider,
    FilterStringRule,
    FilterStringEquals,
    LogicalAndFilter,
    Transaction,
    FamilyInstance
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsInteger,
    get_parameter_value_by_name_AsValueString,
    get_parameter_value_by_name_AsDouble
)
from Parameters.Add_SharedParameters import Shared_Params

import re
import os
from math import atan2
from fractions import Fraction

Shared_Params()

path, filename = os.path.split(__file__)
family_pathCC = os.path.join(path, 'Pipe Label.rfa')

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

FamilyName = 'Pipe Label'
FamilyType = 'Pipe Label'

try:
    level = curview.GenLevel
except:
    level = None


def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        pass


# --------------------------------------------------
# Robust family loading
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


class PipeLabelFamilyManager(object):
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


family_manager = PipeLabelFamilyManager(doc, FamilyName, family_pathCC)
famsymb = family_manager.get_ready_symbol(FamilyType)

if not famsymb:
    raise Exception("Family symbol '{}' was not found or could not be activated.".format(FamilyType))


def set_parameter_by_name_local(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)


def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()


def get_parameter_value_by_name_AsDouble_local(element, parameterName):
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
    if not result:
        raise Exception("Could not project point onto pipe centerline.")
    return result.XYZPoint


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

    def frac2string(s):
        i, f = s.groups(0)
        f = Fraction(f)
        return str(int(i) + float(f))

    overall_size = get_parameter_value_by_name_AsString(pipe, 'Overall Size')
    if not overall_size:
        raise Exception("Pipe does not have a valid Overall Size.")

    if '/' in overall_size:
        diameter = float(
            re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)[^\d.]', frac2string, overall_size)
        ) / 12.0
    else:
        diameter = float(re.sub(r'[^\d.]', '', overall_size)) / 12.0

    set_parameter_by_name_local(new_family_instance, "Diameter", diameter)

    pipe_connectors = list(pipe.ConnectorManager.Connectors)
    if len(pipe_connectors) < 2:
        raise Exception("Selected pipe does not have enough connectors.")

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
    axis = DB.Line.CreateBound(
        insertion_point,
        DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1.0)
    )

    DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)

    set_parameter_by_name_local(
        new_family_instance,
        'FP_Service Name',
        get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Name')
    )
    set_parameter_by_name_local(
        new_family_instance,
        'FP_Service Abbreviation',
        get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Abbreviation')
    )

    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly and level:
        schedule_level_param.Set(level.Id)


def set_pipe_label_size():
    pipe_accessory_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory)
    family_instance_filter = ElementClassFilter(FamilyInstance)

    provider = ParameterValueProvider(ElementId(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM))

    if RevitINT > 2022:
        rule = FilterStringRule(provider, FilterStringEquals(), "Pipe Label")
    else:
        rule = FilterStringRule(provider, FilterStringEquals(), "Pipe Label", False)

    family_name_filter = ElementParameterFilter(rule)

    filter = LogicalAndFilter(pipe_accessory_filter, family_name_filter)

    pipe_labels = FilteredElementCollector(doc, curview.Id) \
        .WherePasses(filter) \
        .WhereElementIsNotElementType() \
        .ToElements()

    t = Transaction(doc, "Set Pipe Label type")
    t.Start()

    for pipe_label in pipe_labels:
        try:
            diameter = get_parameter_value_by_name_AsDouble_local(pipe_label, 'Diameter') * 12
            product_entry = ""

            if diameter <= 0.50:
                product_entry = "AA"
            elif 0.50 < diameter < 1.001:
                product_entry = "A"
            elif 1.00 < diameter < 2.376:
                product_entry = "B"
            elif 2.375 < diameter < 3.251:
                product_entry = "C"
            elif 3.25 < diameter < 4.501:
                product_entry = "D"
            elif 4.50 < diameter < 5.876:
                product_entry = "E"
            elif 5.875 < diameter < 7.876:
                product_entry = "F"
            elif 7.875 < diameter < 9.876:
                product_entry = "G"
            else:
                product_entry = "H"

            product_entry_str = "{:.3f} - {}".format(diameter, product_entry)
            set_parameter_by_name_local(pipe_label, 'FP_Product Entry', product_entry_str)

        except:
            pass

    t.Commit()


while True:
    t = None
    try:
        pipe = select_fabrication_pipe()

        t = Transaction(doc, 'Place Pipe Label Family')
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

set_pipe_label_size()