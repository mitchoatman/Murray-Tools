from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import BoundingBoxXYZ, FilteredElementCollector, Transaction, BuiltInCategory, FamilyInstance
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

def GetCenterPoint(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    center = (bBox.Max + bBox.Min) / 2
    return (center.X, center.Y, center.Z)

def IsNestedFamily(element):
    # Check if the element is a FamilyInstance and if it has a parent (indicating it's nested)
    if isinstance(element, FamilyInstance):
        return element.SuperComponent is not None
    return False

# Create a FilteredElementCollector to get all PipeAccessory elements
AllElements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PlumbingFixtures)\
                   .WhereElementIsNotElementType() \
                   .ToElements()

# Filter out nested families
main_families = [el for el in AllElements if not IsNestedFamily(el)]

# Get the center point of each selected element
element_ids = []
center_points = []

for reference in main_families:
    center_point = GetCenterPoint(reference.Id)
    center_points.append(center_point)
    element_ids.append(reference.Id)

# Find the duplicates in the list of center points
duplicates = []
duplicate_element_ids = []
unique_center_points = []

for i, cp in enumerate(center_points):
    if cp not in unique_center_points:
        unique_center_points.append(cp)
    else:
        duplicates.append(cp)
        duplicate_element_ids.append(element_ids[i])

try:
    if duplicates:
        forms.alert_ifnot(len(duplicates) < 0,
                          ("Delete Duplicate(s): {}".format(len(duplicates))),
                          yes=True, no=True, exitscript=True)
        
        with Transaction(doc, "Delete Elements") as transaction:
            transaction.Start()
            for element_id in duplicate_element_ids:
                doc.Delete(element_id)
            transaction.Commit()
    else:
        forms.show_balloon('Duplicates', 'No Duplicates Found')

except:
    pass
