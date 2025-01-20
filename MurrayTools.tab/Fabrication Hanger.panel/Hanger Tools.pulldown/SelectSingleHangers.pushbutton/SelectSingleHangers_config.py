
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart, FabricationAncillaryUsage, Transaction, TransactionGroup, FilteredElementCollector, ElementCategoryFilter, BuiltInCategory
from Autodesk.Revit.UI.Selection import *
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView



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




