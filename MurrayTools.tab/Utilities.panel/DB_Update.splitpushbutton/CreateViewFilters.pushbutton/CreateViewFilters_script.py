import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from random import randint
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.Attributes import *
from System.Collections.Generic import List
from collections import OrderedDict
from SharedParam.Add_Parameters import Shared_Params
import System
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsValueString, get_parameter_value_by_name_AsInteger

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
curview = doc.ActiveView

# Create a FilteredElementCollector to get all FabricationPart elements from the current view
part_collector = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

SNamelist = []

if part_collector:
    try:
        SNamelist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'), part_collector))
    except:
        print('No Fabrication Parts in View')

SNamelist_set = set(SNamelist)

# 2. create list of categories that will be used for the filter
categories = List[ElementId]()
categories.Add(ElementId(BuiltInCategory.OST_FabricationHangers))
categories.Add(ElementId(BuiltInCategory.OST_FabricationPipework))

# 3a. create rules and filters for each service name
fabrication_service_name_parameter = ElementId(BuiltInParameter.FABRICATION_SERVICE_NAME)

# Collect existing ParameterFilterElements
existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
existing_filter_names = {filter.Name for filter in existing_filters}
existing_filter_dict = {filter.Name: filter.Id for filter in existing_filters}

def random_color():
    r = randint(0, 230)
    g = randint(0, 230)
    b = randint(0, 230)
    return r, g, b

# Looks for Dashed line pattern to collect it's ID.
# Get the line pattern by name
line_patterns = FilteredElementCollector(doc).OfClass(LinePatternElement)
dashed_pattern_id = None
for pattern in line_patterns:
    if pattern.Name == "Dash":
        dashed_pattern_id = pattern.Id
        break

# Check if the dashed line pattern was found
if dashed_pattern_id is None:
    raise ValueError("Dashed line pattern not found in the document.")

# Define a dictionary to map system names to their RGB values
system_colors = {
    # waters
    'IRRIGATION WATER': (170, 191, 255),
    'DOMESTIC COLD WATER': (0, 63, 255),
    'DOMESTIC HOT WATER': (227, 34, 143),
    'HEATING WATER': (227, 148, 20),
    'UG DOMESTIC COLD WATER': (0, 63, 255),
    'CHILLED WATER': (133, 255, 190),
    # waste and vent
    'ACID VENT': (255, 0, 127),
    'SUMP PUMP DISCHARGE': (115, 117, 8),
    'TRAP PRIMER': (12, 38, 207),
    'OVERFLOW DRAIN': (179, 16, 146),
    'SANITARY VENT': (21, 237, 50),
    'SANITARY WASTE': (189, 0, 189),
    'CONDENSATE DRAIN': (86, 129, 118),
    'REFRIGERATION HOT GAS': (21, 130, 77),
    'STORM DRAIN': (157, 56, 224),
    'EMERGENCY DRAIN': (80, 4, 92),
    'GREASE WASTE': (189, 189, 126),
    'GREY WASTE': (138, 109, 58),
    'LAB WASTE': (55, 120, 81),
    'LAB VENT': (83, 184, 123),
    'UG OVERFLOW DRAIN': (179, 16, 146),
    'UG CONDENSATE DRAIN': (26, 112, 66),
    'UG SANITARY VENT': (21, 237, 50),
    'UG SANTIARY WASTE': (153, 5, 143),
    'UG GREY WASTE': (138, 109, 58),
    'UG STORM DRAIN': (157, 56, 224),
    'UG LAB WASTE': (55, 120, 81),
    'UG LAB VENT': (83, 184, 123),
    'UG TRAP PRIMER': (12, 38, 207),
    # gasses
    'CARBON DIOXIDE': (22, 107, 37),
    'LAB AIR': (43, 102, 67),
    'COMPRESSED AIR': (105, 110, 13),
    'GAS': (0, 129, 0),
    'NITROGEN': (41, 5, 66),
    'OXYGEN': (6, 59, 2),
    'UG CARBON DIOXIDE': (22, 107, 37),

    # Add more system names and their corresponding RGB values here
}

# Define a dictionary to store custom filters (OrderedDict to preserve order)
custom_filters = OrderedDict()

# Adding items in the desired order
custom_filters["SPOOL 1"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "1",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (255, 0, 0)  # Red color for visibility
}
custom_filters["SPOOL 2"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "2",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (79, 0, 59)  # White color for visibility
}
custom_filters["SPOOL 3"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "3",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (0, 0, 255)  # Blue color for visibility
}
custom_filters["SPOOL 4"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "4",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (0, 255, 0)  # Green color for visibility
}
custom_filters["SPOOL 5"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "5",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (255, 0, 0)  # Red color for visibility
}
custom_filters["SPOOL 6"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "6",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (79, 0, 59)  # White color for visibility
}
custom_filters["SPOOL 7"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "7",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (0, 0, 255)  # Blue color for visibility
}
custom_filters["SPOOL 8"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "8",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (0, 255, 0)  # Green color for visibility
}
custom_filters["SPOOL 9"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "9",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (255, 0, 0)  # Red color for visibility
}
custom_filters["SPOOL 0"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "0",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers],
    "color": (79, 0, 59)  # White color for visibility
}

# Add more custom filters here if needed



# Get filters already applied to the view
view_template_id = curview.ViewTemplateId
if not view_template_id.Equals(ElementId.InvalidElementId):
    view_to_modify = doc.GetElement(view_template_id)
else:
    view_to_modify = curview

applied_filters = {doc.GetElement(id).Name: id for id in view_to_modify.GetFilters()}

with Transaction(doc, "Create and Apply Filters") as t:
    t.Start()

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

            if condition == "EndsWith":
                rule = ParameterFilterRuleFactory.CreateEndsWithRule(param_id, value, False)
            # Add more conditions as needed
            else:
                raise Exception('Condition not supported')

            filter_element = ElementParameterFilter(rule)
            filter_elem = ParameterFilterElement.Create(doc, filter_name, categories)
            filter_elem.SetElementFilter(filter_element)
            filter_id = filter_elem.Id

        # Check if the filter is already applied to the view
        if filter_name not in applied_filters:
            overrides = OverrideGraphicSettings().SetProjectionLineColor(color)
            view_to_modify.AddFilter(filter_id)
            view_to_modify.SetFilterVisibility(filter_id, True)
            view_to_modify.SetFilterOverrides(filter_id, overrides)

    for service_name in SNamelist_set:
        if service_name in system_colors:
            r, g, b = system_colors[service_name]
        else:
            r, g, b = random_color()

        if service_name in existing_filter_names:
            paramFilterId = existing_filter_dict[service_name]
        else:
            rule = ParameterFilterRuleFactory.CreateEqualsRule(fabrication_service_name_parameter, service_name, False)
            filter = ElementParameterFilter(rule)
            paramFilter = ParameterFilterElement.Create(doc, service_name, categories)
            paramFilter.SetElementFilter(filter)
            paramFilterId = paramFilter.Id

        if service_name not in applied_filters:
            overrides = OverrideGraphicSettings()
            overrides.SetProjectionLineColor(Color(r, g, b))
            #overrides.SetSurfaceTransparency(50)
            #overrides.SetProjectionLinePatternId(dashed_pattern_id)
            #overrides.SetHalftone(True)
            #overrides.SetCutLineWeight(5)

            view_to_modify.AddFilter(paramFilterId)
            view_to_modify.SetFilterVisibility(paramFilterId, True)
            view_to_modify.SetFilterOverrides(paramFilterId, overrides)

    t.Commit()
