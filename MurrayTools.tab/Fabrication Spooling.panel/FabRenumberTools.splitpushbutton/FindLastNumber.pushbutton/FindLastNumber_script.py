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

def get_format_and_number(item_number):
    # This function separates the format and the numeric part of the item number.
    format_part = ''.join(char for char in item_number if not char.isdigit())
    number_part = ''.join(char for char in item_number if char.isdigit())
    # If number_part is an empty string, set it to a default value (e.g., -1)
    number_part = int(number_part) if number_part else -1
    return format_part, number_part


# Initialize a dictionary to store the maximum value for each format
max_values = {}

# Iterate over all elements in the view
for element in part_collector:
    # Retrieve the "Item Number" parameter
    item_number = element.LookupParameter('Item Number').AsString()

    # Separate the format and numeric part of the item number
    format_part, number_part = get_format_and_number(item_number)

    # Update the maximum value for this format if necessary
    if format_part not in max_values or number_part > max_values[format_part]:
        max_values[format_part] = number_part

# Print out the maximum value for each format
for format_part, max_value in max_values.items():
    print("The maximum value for format {}{}".format(format_part, max_value))

