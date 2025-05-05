from Autodesk.Revit import DB
from Autodesk.Revit.DB import Document, Parameter, BuiltInParameter, ElementId
from pyrevit import revit
import sys

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Get currently selected elements
selection = revit.get_selection()

# Filter selection for MEP Fabrication Pipework with CID 2041
Fpipework = []

# Iterate over selection and filter for fabrication pipes with CID 2041
for elem in selection:
    if (elem.Category and elem.Category.Name == "MEP Fabrication Pipework" and 
        hasattr(elem, "ItemCustomId") and elem.ItemCustomId == 2041):
        Fpipework.append(elem)

if len(Fpipework) > 0:
    # Iterate over fabrication pipes and collect length data
    Total_Length = 0.0
    straight_pipes = 0

    for pipe in Fpipework:
        try:
            # Get element type and family name for debugging
            elem_type = doc.GetElement(pipe.GetTypeId())
            family_name = elem_type.FamilyName if elem_type else "Unknown"

            # Check if the element is a straight pipe
            if hasattr(pipe, "IsAStraight") and pipe.IsAStraight:
                len_param = pipe.Parameter[BuiltInParameter.FABRICATION_PART_LENGTH]
                if len_param and len_param.HasValue:
                    length = len_param.AsDouble()
                    Total_Length += length
                    straight_pipes += 1
                    # # Print detailed information for debugging
                    # print("Straight Pipe - ID: {}, Family: {}, CID: {}, Length: {:.2f} ft, IsAStraight: {}".format(
                        # pipe.Id, family_name, pipe.ItemCustomId, length, pipe.IsAStraight))
                # else:
                    # print("Warning: Pipe ID {} has no valid length parameter (Family: {}, CID: {})".format(
                        # pipe.Id, family_name, pipe.ItemCustomId))
            else:
                print("Skipped Element - ID: {}, Family: {}, CID: {}, IsAStraight: {}, Reason: Not straight".format(
                    pipe.Id, family_name, pipe.ItemCustomId, 
                    pipe.IsAStraight if hasattr(pipe, "IsAStraight") else "N/A"))
        except Exception as e:
            print("Error processing Element ID {}: {}".format(pipe.Id, str(e)))

    # Print the total
    print("Total straight pipes processed (CID 2041): {}".format(straight_pipes))
    print("Linear feet of selected straight pipe(s) is: {:.2f}".format(Total_Length))
else:
    print('At least one fabrication pipe with CID 2041 must be selected.')
    sys.exit()