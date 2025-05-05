import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FabricationPart

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Collect all FabricationPart elements in the current view
part_collector = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

def get_format_and_number(item_numbers):
    prefix_max_number = {}

    # Iterate through each item number
    for item_number in item_numbers:
        # Skip if item_number is None or empty
        if not item_number:
            continue
            
        # Check if the item number is purely numeric (no prefix)
        if item_number.isdigit():
            format_part = ''
        else:
            # Iterate through the string from the right side
            for i in range(len(item_number) - 1, -1, -1):
                if not item_number[i].isdigit():
                    # Found a non-digit character, split the string
                    format_part = item_number[:i+1]
                    break
            else:
                # If no non-digit character is found, the format part is empty
                format_part = ''
        
        # Get the numeric part (with or without prefix)
        number_part = item_number[len(format_part):]

        # Skip if number_part is empty or not numeric
        if not number_part or not number_part.isdigit():
            continue

        # Update the maximum number for the prefix
        if format_part not in prefix_max_number:
            prefix_max_number[format_part] = number_part
        else:
            # Check which number is greater considering the padding
            prefix_max_number[format_part] = max(
                prefix_max_number[format_part], 
                number_part, 
                key=lambda x: (len(x), x)
            )

    return prefix_max_number

# Collect FP_Valve Number values
fp_valve_numbers = []

# Iterate over all elements in the view
for element in part_collector:
    # Retrieve the "FP_Valve Number" parameter
    param = element.LookupParameter('FP_Valve Number')
    if param and param.HasValue:
        valve_number = param.AsString()
        if valve_number:  # Ensure the value is not None or empty
            fp_valve_numbers.append(valve_number)

# Get the maximum valve number for each prefix
result = get_format_and_number(fp_valve_numbers)

# Print the results
if result:
    for prefix, max_number in result.items():
        print("The last FP_Valve Number for prefix '{}': {}{}".format(prefix, prefix, max_number))
else:
    print("No valid FP_Valve Numbers found in the current view.")