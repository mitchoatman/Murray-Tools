from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import BoundingBoxXYZ, FilteredElementCollector, Transaction, BuiltInCategory, FabricationPart
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

def GetCenterPoint(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    center = (bBox.Max + bBox.Min) / 2
    return (center.X, center.Y, center.Z)

# Create a FilteredElementCollector to get all PipeAccessory elements
AllElements = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PipeAccessory)\
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
for cp in center_points:
    if cp not in unique_center_points:
        unique_center_points.append(cp)
    else:
        duplicates.append(cp)

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


