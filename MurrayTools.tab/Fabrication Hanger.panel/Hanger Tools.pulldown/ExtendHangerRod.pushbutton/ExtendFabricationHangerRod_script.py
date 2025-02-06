# Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, ReferencePlane

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Selection Filter for Fabrication Hangers
class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, category_name):
        self.category_name = category_name

    def AllowElement(self, e):
        return e.Category and e.Category.Name == self.category_name

    def AllowReference(self, ref, point):
        return False  # Only allow element selection

# Select Fabrication Hangers
pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
                                      CustomISelectionFilter("MEP Fabrication Hangers"),
                                      "Select Fabrication Hangers to Extend")
Hanger = [doc.GetElement(elId) for elId in pipesel]

# Select a Reference Plane
class ReferencePlaneSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, ReferencePlane)  # Only allow Reference Planes

    def AllowReference(self, ref, point):
        return False

ref_plane_ref = uidoc.Selection.PickObject(ObjectType.Element, ReferencePlaneSelectionFilter(),
                                           "Select a reference plane")
ref_plane = doc.GetElement(ref_plane_ref.ElementId)

# Get the plane geometry
plane = ref_plane.GetPlane()
plane_normal = plane.Normal.Normalize()  # Ensure the normal is a unit vector
plane_origin = plane.Origin

# Start transaction
t = Transaction(doc, 'Extend Hanger Rods')
t.Start()

for e in Hanger:
    rod_info = e.GetRodInfo()
    rod_count = rod_info.RodCount
    rod_info.CanRodsBeHosted = False  # Detach rods from structure

    for n in range(rod_count):
        rod_len = rod_info.GetRodLength(n)
        rod_pos = rod_info.GetRodEndPosition(n)  # Should already be in Project Coordinates

        # Calculate intersection of rod with reference plane
        rod_vector = rod_pos - plane_origin
        distance_to_plane = plane_normal.DotProduct(rod_vector)
        intersection_point = rod_pos - (distance_to_plane * plane_normal)

        # Calculate new rod length
        delta_length = intersection_point.DistanceTo(rod_pos)

        # Check if the reference plane is above or below the rod
        if distance_to_plane > 0:
            new_length = rod_len - delta_length
        else:
            new_length = rod_len + delta_length

        # Set the new rod length
        rod_info.SetRodLength(n, new_length)

t.Commit()
