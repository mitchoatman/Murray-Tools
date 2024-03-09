from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import BoundingBoxXYZ, FilteredElementCollector, Transaction, BuiltInCategory, FabricationPart
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def GetCenterPoint(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    center = (bBox.Max + bBox.Min) / 2
    return (center.X, center.Y, center.Z)

# Create a FilteredElementCollector to get all PlumbingFixtures elements
AllElements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PlumbingFixtures) \
                   .WhereElementIsNotElementType() \
                   .ToElements()
# Get the center point of each selected element
element_ids = []
center_points = []

for reference in AllElements:
    center_point = GetCenterPoint(reference.Id)
    center_points.append(center_point)
    element_ids.append(reference.Id)

# Find the duplicates in the list of center points
duplicates = []
unique_center_points = []
for center_point in center_points:
    if center_point not in unique_center_points:
        unique_center_points.append(center_point)
    else:
        duplicates.append(center_point)

# Delete the elements that belong to duplicate center points
try:
    if duplicates:
        forms.alert_ifnot(duplicates < 0,
                          ("Delete Duplicate(s): {}".format(len(duplicates))),
                          yes=True, no=True, exitscript=True)
        
        for duplicate in duplicates:
            duplicate_index = center_points.index(duplicate)
            element_id = element_ids[duplicate_index]
            with Transaction(doc, "Delete Element") as transaction:
                transaction.Start()
                doc.Delete(element_id)
                transaction.Commit()
    else:
        forms.toast(
            'No Duplicates Found',
            title="Duplicates",
            appid="Murray Tools",
            icon="C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\MURRAY RIBBON\Murray.extension\Murray.ico",
            click="https://murraycompany.com",)
except:
    pass


