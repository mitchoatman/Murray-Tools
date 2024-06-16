from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import BoundingBoxXYZ, FilteredElementCollector, Transaction, BuiltInCategory, FabricationPart
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def GetCenterPoint(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    center = (bBox.Max + bBox.Min) / 2
    return (center.X, center.Y, center.Z)

# Create a FilteredElementCollector to get all PipeAccessory elements
AllElements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory) \
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
        forms.toast(
            'No Duplicates Found',
            title="Duplicates",
            appid="Murray Tools",
            icon="",
            click="https://murraycompany.com",)
except:
    pass