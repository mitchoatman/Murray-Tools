import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import FabricationPart, FilteredElementCollector, ParameterFilterRuleFactory, Transaction, Color, LinePatternElement, BuiltInParameter, BuiltInCategory, ElementId, ParameterFilterElement, FilterInverseRule, ElementParameterFilter, OverrideGraphicSettings
from Autodesk.Revit.UI import UIApplication
from System.Collections.Generic import List
from collections import OrderedDict
from random import randint
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

# Create a FilteredElementCollector to get all FabricationPart elements from the current view
part_collector = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

SNamelist = []

selection = [doc.GetElement(id) for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]

# Filter selection for elements that have the "Fabrication Service" parameter
filtered_selection = [x for x in selection if x.LookupParameter("Fabrication Service")]

if filtered_selection:
    try:
        SNamelist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'), filtered_selection))
    except:
        pass
else:
    try:
        SNamelist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'), part_collector))
    except:
        print('No Fabrication Parts in View')

# 2. create list of categories that will be used for the filter
categories = List[ElementId]()
categories.Add(ElementId(BuiltInCategory.OST_FabricationHangers))
categories.Add(ElementId(BuiltInCategory.OST_FabricationPipework))
categories.Add(ElementId(BuiltInCategory.OST_FabricationDuctwork))

# Collect existing ParameterFilterElements
existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
existing_filter_names = {filter.Name for filter in existing_filters}
existing_filter_dict = {filter.Name: filter.Id for filter in existing_filters}

# Define a dictionary to store custom filters (OrderedDict to preserve order)
custom_filters = OrderedDict()
# Adding SPOOL filters
custom_filters["SPOOL 1"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "1",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (255, 0, 0)  # Red color for visibility
}
custom_filters["SPOOL 2"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "2",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (79, 0, 59)  # White color for visibility
}
custom_filters["SPOOL 3"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "3",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 0, 255)  # Blue color for visibility
}
custom_filters["SPOOL 4"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "4",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 255, 0)  # Green color for visibility
}
custom_filters["SPOOL 5"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "5",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (255, 0, 0)  # Red color for visibility
}
custom_filters["SPOOL 6"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "6",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (79, 0, 59)  # White color for visibility
}
custom_filters["SPOOL 7"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "7",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 0, 255)  # Blue color for visibility
}
custom_filters["SPOOL 8"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "8",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 255, 0)  # Green color for visibility
}
custom_filters["SPOOL 9"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "9",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (255, 0, 0)  # Red color for visibility
}
custom_filters["SPOOL 0"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "0",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (79, 0, 59)  # White color for visibility
}

# Get filters already applied to the view
view_template_id = curview.ViewTemplateId
if not view_template_id.Equals(ElementId.InvalidElementId):
    view_to_modify = doc.GetElement(view_template_id)
else:
    view_to_modify = curview

applied_filters = {doc.GetElement(id).Name: id for id in view_to_modify.GetFilters()}

with Transaction(doc, "Create and Apply Filters") as t:
    t.Start()

    # Only apply custom filters when no elements are selected
    if not filtered_selection:
        # Apply custom filters from the dictionary in order
        for filter_name, filter_props in custom_filters.items():
            param_name = filter_props["parameter_name"]
            condition = filter_props["condition"]
            value = filter_props["value"]
            categories = List[ElementId]([ElementId(cat) for cat in filter_props["categories"]])
            color = Color(*filter_props["color"])

            # Check if the filter already exists
            if filter_name in existing_filter_names:
                filter_id = existing_filter_dict[filter_name]
            else:
                # Get the parameter ID from an element that contains the parameter
                sample_element = part_collector[0] if part_collector else None
                if not sample_element:
                    raise Exception('No elements found to retrieve parameter ID from')

                param_id = None
                for p in sample_element.Parameters:
                    if p.Definition.Name == param_name:
                        param_id = p.Id
                        break
                if not param_id:
                    raise Exception('STRATUS Assembly Parameter not found')
                if RevitINT < 2023:
                    if condition == "EndsWith":
                        rule = ParameterFilterRuleFactory.CreateEndsWithRule(param_id, value, False)
                    else:
                        raise Exception('Condition not supported')
                else:
                    if condition == "EndsWith":
                        rule = ParameterFilterRuleFactory.CreateEndsWithRule(param_id, value)
                    else:
                        raise Exception('Condition not supported')

                filter_element = ElementParameterFilter(rule)
                filter_elem = ParameterFilterElement.Create(doc, filter_name, categories)
                filter_elem.SetElementFilter(filter_element)
                filter_id = filter_elem.Id

                # Check if the filter is already applied to the view
                if filter_name not in applied_filters:
                    overrides = OverrideGraphicSettings()
                    overrides.SetProjectionLineColor(color)
                    view_to_modify.AddFilter(filter_id)
                    view_to_modify.SetFilterVisibility(filter_id, True)
                    view_to_modify.SetFilterOverrides(filter_id, overrides)

    t.Commit()