import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, FabricationConfiguration, FabricationPart

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

part_collector = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

def get_format_and_number(item_numbers):
    prefix_max_number = {}

    # Iterate through each item number
    for item_number in item_numbers:
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

item_numbers = []

# Iterate over all elements in the view
for element in part_collector:
    # Retrieve the "Item Number" parameter
    item_number = element.LookupParameter('Item Number').AsString()
    item_numbers.append(element.LookupParameter('Item Number').AsString())
result = get_format_and_number(item_numbers)

for prefix, max_number in result.items():
    print("The maximum value for format:  {}{}".format(prefix, max_number))




