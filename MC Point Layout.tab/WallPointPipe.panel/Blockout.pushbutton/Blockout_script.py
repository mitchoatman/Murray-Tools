from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, Transaction, XYZ
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString
from Autodesk.Revit.UI import TaskDialog
from math import atan2
import os
import clr

from Parameters.Add_SharedParameters import Shared_Params
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# --------------------------------------------------
# Selection filter
# --------------------------------------------------
class PickByCategorySelectionFilter(ISelectionFilter):
    def __init__(self, category_ids):
        self.category_ids = category_ids

    def AllowElement(self, element):
        if element.Category and element.Category.Id in self.category_ids:
            return True
        return False

    def AllowReference(self, reference, point):
        return False


def select_fabrication_parts():
    try:
        category_ids = [
            DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_FabricationPipework).Id
        ]
        msfilter = PickByCategorySelectionFilter(category_ids)
        selection = uidoc.Selection.PickObjects(
            ObjectType.Element,
            msfilter,
            "Select MEP Fabrication Pipework to Blockout"
        )
        return [doc.GetElement(ref.ElementId) for ref in selection]
    except:
        return []


# --------------------------------------------------
# Geometry helpers
# --------------------------------------------------
def get_combined_bounding_box(elements):
    if not elements:
        return None, 0, 0

    first_pipe_bb = elements[0].get_BoundingBox(doc.ActiveView)
    if not first_pipe_bb:
        return None, 0, 0

    delta_x = abs(first_pipe_bb.Max.X - first_pipe_bb.Min.X)
    delta_y = abs(first_pipe_bb.Max.Y - first_pipe_bb.Min.Y)

    combined_min = first_pipe_bb.Min
    combined_max = first_pipe_bb.Max

    insulation_thickness = elements[0].InsulationThickness if hasattr(elements[0], 'InsulationThickness') and elements[0].HasInsulation else 0.0
    if delta_x > delta_y:
        combined_min = XYZ(combined_min.X, combined_min.Y - insulation_thickness, combined_min.Z - insulation_thickness)
        combined_max = XYZ(combined_max.X, combined_max.Y + insulation_thickness, combined_max.Z + insulation_thickness)
    else:
        combined_min = XYZ(combined_min.X - insulation_thickness, combined_min.Y, combined_min.Z - insulation_thickness)
        combined_max = XYZ(combined_max.X + insulation_thickness, combined_max.Y, combined_max.Z + insulation_thickness)

    for element in elements[1:]:
        bb = element.get_BoundingBox(doc.ActiveView)
        if not bb:
            continue

        insulation_thickness = element.InsulationThickness if hasattr(element, 'InsulationThickness') and element.HasInsulation else 0.0
        if delta_x > delta_y:
            element_min = XYZ(bb.Min.X, bb.Min.Y - insulation_thickness, bb.Min.Z - insulation_thickness)
            element_max = XYZ(bb.Max.X, bb.Max.Y + insulation_thickness, bb.Max.Z + insulation_thickness)
        else:
            element_min = XYZ(bb.Min.X - insulation_thickness, bb.Min.Y, bb.Min.Z - insulation_thickness)
            element_max = XYZ(bb.Max.X + insulation_thickness, bb.Max.Y, bb.Max.Z + insulation_thickness)

        combined_min = XYZ(min(combined_min.X, element_min.X), min(combined_min.Y, element_min.Y), min(combined_min.Z, element_min.Z))
        combined_max = XYZ(max(combined_max.X, element_max.X), max(combined_max.Y, element_max.Y), max(combined_max.Z, element_max.Z))

    bbox = DB.BoundingBoxXYZ()
    bbox.Min = combined_min
    bbox.Max = combined_max
    return bbox, delta_x, delta_y


def get_duct_centerline(element):
    pipe_location = element.Location
    if isinstance(pipe_location, DB.LocationCurve):
        return pipe_location.Curve
    raise Exception("The selected element does not have a valid centerline.")


def get_center_of_bounding_box(bbox):
    if not bbox:
        return None
    center_x = (bbox.Min.X + bbox.Max.X) / 2
    center_y = (bbox.Min.Y + bbox.Max.Y) / 2
    center_z = (bbox.Min.Z + bbox.Max.Z) / 2
    return DB.XYZ(center_x, center_y, center_z)


def get_projected_point_on_axis(elements, user_point, bbox, delta_x, delta_y):
    centerline = get_duct_centerline(elements[0])
    start_point = centerline.GetEndPoint(0)
    end_point = centerline.GetEndPoint(1)
    direction = (end_point - start_point).Normalize()

    vector_to_point = user_point - start_point
    projection_length = vector_to_point.DotProduct(direction)

    center = get_center_of_bounding_box(bbox)
    if delta_x > delta_y:
        insertion_point = XYZ(start_point.X + direction.X * projection_length, center.Y, center.Z)
    else:
        insertion_point = XYZ(center.X, start_point.Y + direction.Y * projection_length, center.Z)

    return insertion_point


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


class BlockoutFamilyManager(object):
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
            TaskDialog.Show("Error", "Type '{}' not found in family '{}'.".format(type_name, self.family_name))
            return None

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


# --------------------------------------------------
# Placement
# --------------------------------------------------
def place_and_modify_family(elements, famsymb, annular_space):
    bbox, delta_x, delta_y = get_combined_bounding_box(elements)
    if not bbox:
        raise Exception("Unable to calculate combined bounding box")

    try:
        user_point = uidoc.Selection.PickPoint("Select location for blockout")
    except:
        raise Exception("Point selection cancelled or failed")

    insertion_point = get_projected_point_on_axis(elements, user_point, bbox, delta_x, delta_y)
    level = doc.GetElement(elements[0].LevelId)

    new_family_instance = doc.Create.NewFamilyInstance(
        insertion_point,
        famsymb,
        level,
        DB.Structure.StructuralType.NonStructural
    )

    if delta_x > delta_y:
        width = (bbox.Max.Y - bbox.Min.Y) + annular_space / 12.0
    else:
        width = (bbox.Max.X - bbox.Min.X) + annular_space / 12.0

    height = (bbox.Max.Z - bbox.Min.Z) + annular_space / 12.0

    set_parameter_by_name(new_family_instance, 'Width', width)
    set_parameter_by_name(new_family_instance, 'Height', height)
    set_parameter_by_name(new_family_instance, 'Length', 0.25)

    duct_connectors = list(elements[0].ConnectorManager.Connectors)
    connector1, connector2 = duct_connectors[0], duct_connectors[1]

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
        get_parameter_value_by_name_AsString(elements[0], 'Fabrication Service Name')
    )

    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly:
        schedule_level_param.Set(level.Id)
    else:
        TaskDialog.Show('Warning', "'Schedule Level' parameter not found in family instance")


# --------------------------------------------------
# Family / settings setup
# --------------------------------------------------
path, filename = os.path.split(__file__)
family_filename = 'BLOCKOUT.rfa'
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
    TaskDialog.Show('Warning', "Failed to read annular space, using default 1 inch")

FamilyName = 'BLOCKOUT'
FamilyType = 'BLOCKOUT'

family_manager = BlockoutFamilyManager(doc, FamilyName, family_pathCC)
famsymb = family_manager.get_ready_symbol(FamilyType)


# --------------------------------------------------
# Main
# --------------------------------------------------
if famsymb:
    try:
        elements = select_fabrication_parts()
        if not elements:
            pass
        else:
            t = Transaction(doc, 'Place Trimble Wall Sleeve Family')
            t.Start()
            try:
                place_and_modify_family(elements, famsymb, AnnularSpace)
                t.Commit()
            except Exception as e:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                TaskDialog.Show('Error', "Error during operation: {}".format(e))
    except Exception as e:
        TaskDialog.Show('Error', "Error during operation: {}".format(e))