import sys
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, XYZ, BoundingBoxXYZ, ViewType, Transaction
from Autodesk.Revit.UI import TaskDialog
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Check if active view is a 3D view
curview = doc.ActiveView
if curview.ViewType != ViewType.ThreeD:
    TaskDialog.Show("Error", "The active view must be a 3D view. Please switch to a 3D view and try again.")
    sys.exit()

# Collect all scope boxes in the model
scope_boxes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType().ToElements()

if not scope_boxes:
    TaskDialog.Show("Error", "No scope boxes found in the model.")
    sys.exit()

# Collect elements in the active view for specified categories
hanger_collector = list(FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers).WhereElementIsNotElementType().ToElements())
pipe_collector = list(FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework).WhereElementIsNotElementType().ToElements())
duct_collector = list(FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationDuctwork).WhereElementIsNotElementType().ToElements())
pipe_accessories = list(FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements())
duct_accessories = list(FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_DuctAccessory).WhereElementIsNotElementType().ToElements())
plumbing_fixtures = list(FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PlumbingFixtures).WhereElementIsNotElementType().ToElements())
structural_stiffener = list(FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_StructuralStiffener).WhereElementIsNotElementType().ToElements())

# Combine all elements into a single list
elements_in_view = hanger_collector + pipe_collector + duct_collector + pipe_accessories + duct_accessories + plumbing_fixtures + structural_stiffener

if not elements_in_view:
    TaskDialog.Show("Error", "No Fabrication Hangers, Pipework, Ductwork, Pipe Accessories, Duct Accessories, or Plumbing Fixtures found in the active view.")
    sys.exit()

# Function to check if a point is inside a bounding box
def is_point_inside_bb(point, bb):
    return (bb.Min.X <= point.X <= bb.Max.X and
            bb.Min.Y <= point.Y <= bb.Max.Y and
            bb.Min.Z <= point.Z <= bb.Max.Z)

# Function to get the center point of an element
def get_element_center(element):
    bb = element.get_BoundingBox(None)
    if bb:
        center = XYZ((bb.Min.X + bb.Max.X) / 2,
                     (bb.Min.Y + bb.Max.Y) / 2,
                     (bb.Min.Z + bb.Max.Z) / 2)
        return center
    else:
        loc = element.Location
        if isinstance(loc, Autodesk.Revit.DB.LocationPoint):
            return loc.Point
        elif isinstance(loc, Autodesk.Revit.DB.LocationCurve):
            return loc.Curve.Evaluate(0.5, True)
        else:
            return None

def set_customdata_by_custid(fabpart, custid, value):
    try:
        fabpart.SetPartCustomDataText(custid, value)
    except:
        pass  # Skip if custom data cannot be set

t = None
try:
    t = Transaction(doc, 'Fabrication Location')
    t.Start()

    for element in elements_in_view:
        center = get_element_center(element)
        if not center:
            continue  # Skip elements without a determinable center

        param_exist = element.LookupParameter("FP_Location")
        if not param_exist:
            continue  # Skip elements without the parameter

        assigned = False
        for scope in scope_boxes:
            bb = scope.get_BoundingBox(None)
            if bb and is_point_inside_bb(center, bb):
                set_parameter_by_name(element, "FP_Location", scope.Name)
                if element.LookupParameter("Fabrication Service"):
                    set_customdata_by_custid(element, 13, scope.Name)
                assigned = True
                break  # Assign to the first matching scope box

        if not assigned:
            set_parameter_by_name(element, "FP_Location", "")  # Clear parameter if no scope box match

    t.Commit()
    message = """Location names have been assigned to:
    - Fabrication Hangers
    - Fabrication Pipework
    - Fabrication Ductwork
    - Pipe Accessories
    - Duct Accessories
    - Plumbing Fixtures
    based on proximity to scope boxes."""
    TaskDialog.Show("Success", message)
except Exception as e:
    TaskDialog.Show("Error", "Error: {}".format(str(e)))
    if t is not None and t.HasStarted():
        t.RollBack()