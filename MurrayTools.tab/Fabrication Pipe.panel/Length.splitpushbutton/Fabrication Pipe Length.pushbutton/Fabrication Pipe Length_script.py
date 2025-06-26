import clr
clr.AddReference('System.Windows.Forms')
import System
from System.Windows.Forms import MessageBox
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter, ElementId,
    FabricationPart
)
from Autodesk.Revit.UI.Selection import ObjectType
import sys

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

# Get selected elements or prompt for selection
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    try:
        picked_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select fabrication pipework elements.")
        selected_ids = [ref.ElementId for ref in picked_refs]
    except:
        MessageBox.Show("Selection cancelled. No elements selected.", "Error")
        sys.exit()

if not selected_ids:
    MessageBox.Show("No elements selected. Please select elements and try again.", "Error")
    sys.exit()

# Collect MEP Fabrication Pipework elements with CID 2041 from selected IDs
selected_id_set = set(str(id.IntegerValue) for id in selected_ids)  # Convert to set of string IDs for comparison
collector = (
    FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_FabricationPipework)
    .WhereElementIsNotElementType()
)

Fpipework = [
    elem for elem in collector
    if isinstance(elem, FabricationPart) and elem.ItemCustomId == 2041 and str(elem.Id.IntegerValue) in selected_id_set
]

if not Fpipework:
    print "No fabrication pipes with CID 2041 selected."
    sys.exit()

# Process fabrication pipes
Total_Length = 0.0
straight_pipes = 0
type_cache = {}  # Cache element types

for pipe in Fpipework:
    try:
        # Get element type from cache or document
        type_id = pipe.GetTypeId().IntegerValue
        if type_id not in type_cache:
            type_cache[type_id] = doc.GetElement(pipe.GetTypeId())
        elem_type = type_cache[type_id]
        family_name = elem_type.FamilyName if elem_type else "Unknown"

        # Check if the pipe is straight and has a valid length
        if getattr(pipe, "IsAStraight", False):
            len_param = pipe.Parameter[BuiltInParameter.FABRICATION_PART_LENGTH]
            if len_param and len_param.HasValue:
                Total_Length += len_param.AsDouble()
                straight_pipes += 1
            else:
                print "Warning: Pipe ID {} has no valid length parameter (Family: {}, CID: {})".format(
                    pipe.Id, family_name, pipe.ItemCustomId)
        else:
            print "Skipped Element - ID: {}, Family: {}, CID: {}, IsAStraight: {}, Reason: Not straight".format(
                pipe.Id, family_name, pipe.ItemCustomId, 
                getattr(pipe, "IsAStraight", "N/A")
            )
    except Exception as e:
        print "Error processing Element ID {}: {}".format(pipe.Id, str(e))

# Output results
print "Total straight pipes processed (CID 2041): {}".format(straight_pipes)
print "Linear feet of selected straight pipe(s): {:.2f}".format(Total_Length)