
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart, FabricationAncillaryUsage, Transaction, TransactionGroup, FilteredElementCollector, ElementCategoryFilter, BuiltInCategory
from Autodesk.Revit.UI.Selection import *
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

def convert_fractions(string):
    # Remove double quotes (") from the input string.
    string = string.replace('"', '')
    # Split the string into a list of tokens.
    tokens = string.split()
    # Initialize variables to keep track of the integer and fractional parts.
    integer_part = 0
    fractional_part = 0.0
    # Iterate over the tokens and convert the mixed number to a float.
    for token in tokens:
        if " " in token:
            # Split the mixed number into integer and fractional parts.
            integer_part_str, fractional_part_str = token.split(" ")
            # Convert the integer part to an integer.
            integer_part += int(integer_part_str)
            # Convert the fractional part to a float.
            fractional_part_str = fractional_part_str.replace('/', '')
            fractional_part += float(fractional_part_str)
        elif "/" in token:
            # If the token is just a fraction, convert it to a float and add it to the fractional part.
            numerator, denominator = token.split("/")
            fractional_part += float(numerator) / float(denominator)
        else:
            # If the token is a standalone number, add it to the integer part.
            integer_part += float(token)
    # Calculate the final result by adding the integer and fractional parts together.
    result = integer_part + fractional_part
    return result

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsString()

def get_parameter_value_by_name_AsValueString(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()

# Creating collector instance and collecting all the fabrication hangers from the model
hangers = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()


hangers_without_host_printed = False  # Flag to ensure the message is printed only once
for hanger in hangers:
    hosted_info = hanger.GetHostedInfo().HostId
    try:
        # Get the host element's size
        HostSize = convert_fractions(get_parameter_value_by_name(doc.GetElement(hosted_info), 'Size'))
    except:
        if not hangers_without_host_printed:
            print("HANGERS WITHOUT A HOST")  # Print the message only once
            hangers_without_host_printed = True
        output = script.get_output()
        print('{}: {}'.format((get_parameter_value_by_name_AsValueString(hanger, 'Family')), output.linkify(hanger.Id)))




