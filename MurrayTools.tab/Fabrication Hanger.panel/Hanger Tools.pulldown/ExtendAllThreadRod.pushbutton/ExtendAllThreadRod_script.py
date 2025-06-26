# Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, ReferencePlane

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Selection Filter for MEP Fabrication Pipework + All Thread Rod family
class CustomISelectionFilter(ISelectionFilter):
    def AllowElement(self, e):
        if not e.Category:
            return False
        if e.Category.Name != "MEP Fabrication Pipework":
            return False
        fam_param = e.LookupParameter("Family")
        return fam_param and fam_param.AsValueString() == "All Thread Rod"

    def AllowReference(self, ref, point):
        return False  # Only allow element selection

# Select All Thread Rods
pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
                                      CustomISelectionFilter(),
                                      "Select 'All Thread Rod' Pipework to Extend")
rods = [doc.GetElement(elId) for elId in pipesel]

# Select a Reference Plane
class ReferencePlaneSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, ReferencePlane)

    def AllowReference(self, ref, point):
        return False

ref_plane_ref = uidoc.Selection.PickObject(ObjectType.Element, ReferencePlaneSelectionFilter(),
                                           "Select a reference plane")
ref_plane = doc.GetElement(ref_plane_ref.ElementId)

# Get the plane geometry
plane = ref_plane.GetPlane()
plane_normal = plane.Normal.Normalize()
plane_origin = plane.Origin

# Start transaction
t = Transaction(doc, 'Extend Rods to Reference Plane')
t.Start()

for e in rods:
    # Use bounding box top point as rod endpoint
    bbox = e.get_BoundingBox(None)
    if not bbox:
        continue

    rod_top = bbox.Max
    rod_bot = bbox.Min
    current_length = rod_top.Z - rod_bot.Z

    # Vector from plane to rod top
    vec_to_plane = rod_top - plane_origin
    distance_to_plane = plane_normal.DotProduct(vec_to_plane)
    intersection_point = rod_top - (distance_to_plane * plane_normal)
    delta_length = intersection_point.DistanceTo(rod_top)

    # Extend or shorten based on side of plane
    if distance_to_plane > 0:
        new_length = current_length - delta_length
    else:
        new_length = current_length + delta_length

    # Set new Length param
    length_param = e.LookupParameter("Length")
    if length_param and not length_param.IsReadOnly:
        length_param.Set(new_length)
    else:
        print("Can't set Length on", e.Id)

t.Commit()
