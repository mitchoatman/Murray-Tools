#Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FamilySymbol, Structure, Transaction, BuiltInParameter, \
                                Family, TransactionGroup, FamilyInstance, ReferencePlane

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView


class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return True

try:
    pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers to Extend")
    Hanger = [doc.GetElement(elId) for elId in pipesel]
except:
    Hanger = False
    pass

collector = FilteredElementCollector(doc)
reference_planes = collector.OfCategory(BuiltInCategory.OST_CLines).WhereElementIsNotElementType().ToElements()

# Initialize a dictionary to store reference plane names
reference_plane_names = {}

# Iterate through the reference planes and get their names
for ref_plane in reference_planes:
    if isinstance(ref_plane, ReferencePlane):
        name = ref_plane.Name
        reference_plane_names[ref_plane.Id] = name

if Hanger:
    if len(Hanger) > 0:
        try:
            ref_plane = uidoc.Selection.PickObject(ObjectType.Element, "Select a reference plane")
            ref_plane = doc.GetElement(ref_plane.ElementId)
            if not isinstance(ref_plane, DB.ReferencePlane):
                ref_plane = None

            # Get the plane geometry
            plane = ref_plane.GetPlane()
            plane_normal = plane.Normal
            plane_origin = plane.Origin

            t = Transaction(doc, 'Extend Hanger Rods')
            t.Start()

            for e in Hanger:
                STName = e.GetRodInfo().RodCount
                # Detaches rods from structure
                hgrhost = e.GetRodInfo().CanRodsBeHosted = False
                STName1 = e.GetRodInfo()
                for n in range(STName):
                    rodlen = STName1.GetRodLength(n)
                    rodpos = STName1.GetRodEndPosition(n)
                    
                    # Calculate the intersection of the rod with the reference plane
                    rod_vector = rodpos - plane_origin
                    distance_to_plane = plane_normal.DotProduct(rod_vector)
                    intersection_point = rodpos - (distance_to_plane * plane_normal)
                    
                    # Calculate the new length of the rod
                    delta_length = intersection_point.DistanceTo(rodpos)
                    
                    # Check if the reference plane is above or below the rod
                    if distance_to_plane > 0:
                        new_length = rodlen - delta_length
                    else:
                        new_length = rodlen + delta_length

                    # Set the new rod length
                    STName1.SetRodLength(n, new_length)

            t.Commit()
        except:
            pass


