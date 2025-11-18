import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, Structure, XYZ, LocationCurve, TransactionGroup
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import TaskDialog
import os, re, sys
from math import atan2, pi
from fractions import Fraction
from Parameters.Add_SharedParameters import Shared_Params
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

path, filename = os.path.split(__file__)
NewFilename = r'\Strut Clamp.rfa'

# Family loading options with minimal parameter overwrites
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

# Selection filter for MEP Fabrication Pipework and Hangers
class PickByCategorySelectionFilter(ISelectionFilter):
    """Selection filter implementation"""
    def __init__(self, category_ids):
        self.category_ids = category_ids

    def AllowElement(self, element):
        """Is element allowed to be selected?"""
        if element.Category and element.Category.Id in self.category_ids:
            return True
        return False

    def AllowReference(self, reference, point):
        """Not used for selection"""
        return False

def select_fabrication_elements():
    try:
        category_ids = [
            DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_FabricationPipework).Id,
            DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_FabricationHangers).Id
        ]
        msfilter = PickByCategorySelectionFilter(category_ids)
        selection = uidoc.Selection.PickObjects(ObjectType.Element, msfilter, "Select MEP Fabrication Pipes and Hangers")
        pipes = []
        hangers = []
        for ref in selection:
            element = doc.GetElement(ref.ElementId)
            if element.Category and element.Category.Id == category_ids[0]:  # OST_FabricationPipework
                pipes.append(element)
            elif element.Category and element.Category.Id == category_ids[1]:  # OST_FabricationHangers
                hangers.append(element)
        return pipes, hangers
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        TaskDialog.Show("Selection Cancelled", "Selection was cancelled by the user. Please select at least one MEP Fabrication Pipe or Hanger to continue.")
        sys.exit()  
    except Exception, e:
        TaskDialog.Show("Error", "An unexpected error occurred during selection: " + str(e))
        sys.exit() 

def get_pipe_centerline(pipe):
    location = pipe.Location
    if isinstance(location, LocationCurve):
        return location.Curve
    raise Exception("Selected pipe has no valid centerline")

def get_pipe_size(pipe):
    overall_size = pipe.get_Parameter(DB.BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
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
            diameter = 0.5
    return diameter / 12  # Convert to feet

def get_OD(pipe):
    outside_diameter_param = pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsValueString()
    cleaned_size = re.sub(r'["]', '', outside_diameter_param.strip())
    try:
        rad = float(cleaned_size) / 2

    except ValueError:
        match = re.match(r'(?:(\d+)[-\s])?(\d+/\d+)', cleaned_size)
        if match:
            integer_part, fraction_part = match.groups()
            rad = float(Fraction(fraction_part))
            if integer_part:
                rad += float(integer_part) / 2
        else:
            rad = 0.25
    return rad / 12  # Convert to feet

def get_hanger_origin(hanger):
    origin = hanger.Origin
    if isinstance(origin, DB.XYZ):
        return origin
    raise Exception("Selected hanger has no valid origin point")

def project_point_on_curve(point, curve):
    return curve.Project(point).XYZPoint

def get_pipe_reference_level(pipe):
    level_param = pipe.LookupParameter("Reference Level")
    if level_param and level_param.StorageType == DB.StorageType.ElementId:
        level_id = level_param.AsElementId()
        if RevitINT > 2025:
            if level_id.Value != -1:
                return doc.GetElement(level_id)
        else:
            if level_id.IntegerValue != -1:
                return doc.GetElement(level_id)            
    return None

def get_existing_strut_clamps():
    strut_clamps = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory)\
                                               .WhereElementIsNotElementType()\
                                               .ToElements()
    existing_locations = {}
    for sc in strut_clamps:
        if sc.Symbol.Family.Name == FamilyName:
            loc = sc.Location
            if isinstance(loc, DB.LocationPoint):
                existing_locations[sc.Id] = loc.Point
    return existing_locations

def is_location_occupied(target_point, existing_locations, tolerance=0.01):
    for loc in existing_locations.values():
        if target_point.DistanceTo(loc) <= tolerance:
            return True
    return False

try:
    # Select pipes and hangers
    pipes, hangers = select_fabrication_elements()
    if not pipes:
        raise Exception("No pipes selected")
    if not hangers:
        raise Exception("No hangers selected")

    # Check if family is in project
    FamilyName = 'Strut Clamp'
    families = FilteredElementCollector(doc).OfClass(Family)
    target_family = next((f for f in families if f.Name == FamilyName), None)
    Fam_is_in_project = target_family is not None

    # Get existing Strut Clamp locations
    existing_strut_locations = get_existing_strut_clamps()

    # Start transaction group
    with TransactionGroup(doc, "Place Strut Clamp on Pipes") as tg:
        tg.Start()

        # Load family if not present
        if not Fam_is_in_project:
            t = Transaction(doc, 'Load Strut Clamp Family')
            t.Start()
            fload_handler = FamilyLoaderOptionsHandler()
            family_path = path + NewFilename
            doc.LoadFamily(family_path, fload_handler)
            t.Commit()

        # Get family symbol
        family_symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
        target_famtype = next((fs for fs in family_symbols if fs.Family.Name == FamilyName), None)

        if target_famtype:
            for pipe in pipes:
                pipe_size = get_pipe_size(pipe)
                pipe_od = get_OD(pipe)
                centerline = get_pipe_centerline(pipe)
                pipe_level = get_pipe_reference_level(pipe)

                for hanger in hangers:
                    t = Transaction(doc, 'Place Strut Clamp')
                    t.Start()

                    target_famtype.Activate()
                    doc.Regenerate()

                    hanger_origin = get_hanger_origin(hanger)
                    projected_point = project_point_on_curve(hanger_origin, centerline)
                    insertion_point = DB.XYZ(projected_point.X, projected_point.Y, projected_point.Z)

                    # Check if location is already occupied
                    if not is_location_occupied(insertion_point, existing_strut_locations):
                        new_lrd = doc.Create.NewFamilyInstance(
                            insertion_point,
                            target_famtype,
                            DB.Structure.StructuralType.NonStructural
                        )

                        diameter_param = new_lrd.LookupParameter("Diameter")
                        if diameter_param:
                            diameter_param.Set(pipe_size)

                        rad_param = new_lrd.LookupParameter("rad")
                        if rad_param:
                            rad_param.Set(pipe_od)

                        service_name_param = pipe.LookupParameter("Fabrication Service Name")
                        if service_name_param:
                            fp_service_param = new_lrd.LookupParameter("FP_Service Name")
                            if fp_service_param:
                                fp_service_param.Set(service_name_param.AsString())

                        pipe_connectors = list(pipe.ConnectorManager.Connectors)
                        vec_x = pipe_connectors[1].Origin.X - pipe_connectors[0].Origin.X
                        vec_y = pipe_connectors[1].Origin.Y - pipe_connectors[0].Origin.Y
                        angle = atan2(vec_y, vec_x) + (pi / 2)
                        axis = DB.Line.CreateBound(insertion_point, 
                                                 DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1))
                        DB.ElementTransformUtils.RotateElement(doc, new_lrd.Id, axis, angle)

                        schedule_level_param = new_lrd.LookupParameter("Schedule Level")
                        if schedule_level_param and pipe_level:
                            schedule_level_param.Set(pipe_level.Id)
                        else:
                            TaskDialog.Show("Warning", "Could not set Schedule Level for Strut Clamp on pipe")

                        # Update existing locations to include the new instance
                        existing_strut_locations[new_lrd.Id] = insertion_point

                    t.Commit()
        else:
            TaskDialog.Show("Error", "Could not find Strut Clamp family symbol")

        tg.Assimilate()

except Exception as e:
    TaskDialog.Show("Error", "Error: %s" % str(e))