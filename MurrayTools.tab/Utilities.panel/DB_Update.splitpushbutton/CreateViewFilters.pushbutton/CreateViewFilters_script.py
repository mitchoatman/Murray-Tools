import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import FabricationPart, FilteredElementCollector, ParameterFilterRuleFactory, Transaction, Color, LinePatternElement, BuiltInParameter, BuiltInCategory, ElementId, ParameterFilterElement \
                                , FilterInverseRule, ElementParameterFilter, OverrideGraphicSettings
from Autodesk.Revit.UI import UIApplication
from System.Collections.Generic import List
from collections import OrderedDict
from random import randint
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString

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

SNamelist_set = set(SNamelist)

# 2. create list of categories that will be used for the filter
categories = List[ElementId]()
categories.Add(ElementId(BuiltInCategory.OST_FabricationHangers))
categories.Add(ElementId(BuiltInCategory.OST_FabricationPipework))
categories.Add(ElementId(BuiltInCategory.OST_FabricationDuctwork))

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
    # mech pipe
    'HEATING HOT WATER RETURN': (255, 127, 0),
    'HEATING HOT WATER SUPPLY': (255, 0, 127),
    'CHILLED WATER RETURN': (127, 255, 191),
    'CHILLED WATER SUPPLY': (255, 191, 0),
    'CONDENSER WATER RETURN': (0, 191, 255),
    'CONDENSER WATER SUPPLY': (0, 63, 255),
    'REFRIGERATION HOT GAS': (21, 130, 77),
    'REFRIGERATION LIQUID': (0, 204, 102),
    'REFRIGERATION SUCTION': (95, 76, 153),
    'CLEAN STEAM': (63, 0, 127),
    'HIGH PRESSURE STEAM': (255, 255, 127),
    'LOW PRESSURE STEAM': (0, 153, 0),
    'MEDIUM PRESSURE STEAM': (255, 127, 255),
    'HIGH PRESSURE CONDENSATE': (76, 153, 133),
    'LOW PRESSURE CONDENSATE': (255, 0, 255),
    'MEDIUM PRESSURE CONDENSATE': (255, 127, 127),
    'UG HEATING HOT WATER RETURN': (255, 127, 0),
    'UG HEATING HOT WATER SUPPLY': (255, 0, 127),
    'UG CHILLED WATER RETURN': (127, 255, 191),
    'UG CHILLED WATER SUPPLY': (255, 191, 0),
    'UG CONDENSER WATER RETURN': (0, 191, 255),
    'UG CONDENSER WATER SUPPLY': (0, 63, 255),

    # waters
    'IRRIGATION WATER': (170, 191, 255),
    'INDUSTRIAL COLD WATER': (0, 191, 255),
    'DOMESTIC COLD WATER': (0, 63, 255),
    'DOMESTIC HOT WATER': (227, 34, 143),
    'DOMESTIC HOT WATER RETURN': (255, 127, 0),
    'SOFT COLD WATER': (255, 255, 127),
    'DE-IONIZED WATER SUPPLY': (255, 0, 127),
    'DE-IONIZED WATER RETURN': (204, 102, 153),
    'UG DOMESTIC COLD WATER': (0, 63, 255),
    'UG DOMESTIC HOT WATER': (227, 34, 143),

    # waste and vent
    'ACID VENT': (255, 0, 127),
    'SUMP PUMP DISCHARGE': (115, 117, 8),
    'TRAP PRIMER': (12, 38, 207),
    'OVERFLOW DRAIN': (204, 0, 102),
    'STORM DRAIN': (255, 127, 191),
    'SANITARY VENT': (21, 237, 50),
    'SANITARY WASTE': (204, 0, 204),
    'CONDENSATE DRAIN': (86, 129, 118),
    'EMERGENCY DRAIN': (80, 4, 92),
    'GREASE WASTE': (189, 189, 126),
    'GREY WASTE': (204, 153, 102),
    'LAB WASTE': (55, 120, 81),
    'LAB VENT': (83, 184, 123),
    'METHANE VENT': (255, 191, 0),
    'UG OVERFLOW DRAIN': (204, 0, 102),
    'UG CONDENSATE DRAIN': (86, 129, 118),
    'UG SANITARY VENT': (21, 237, 50),
    'UG SANTIARY WASTE': (204, 0, 204),
    'UG GREY WASTE': (204, 153, 102),
    'UG STORM DRAIN': (255, 127, 191),
    'UG LAB WASTE': (55, 120, 81),
    'UG LAB VENT': (83, 184, 123),
    'UG TRAP PRIMER': (12, 38, 207),
    'UG GREASE WASTE': (189, 189, 126),
    'FUEL OIL SUPPLY': (175, 175, 000),

    # gasses
    'CARBON DIOXIDE': (22, 107, 37),
    'LAB AIR': (76, 153, 133),
    'LAB VACUUM': (255, 0, 191),
    'COMPRESSED AIR': (105, 110, 13),
    'MEDICAL AIR': (76, 153, 133),
    'MEDICAL VACUUM': (255, 0, 191),
    'GAS': (0, 129, 0),
    'NITROGEN': (41, 5, 66),
    'OXYGEN': (6, 59, 2),
    'UG NITROGEN': (41, 5, 66),
    'UG OXYGEN': (6, 59, 2),
    'UG CARBON DIOXIDE': (22, 107, 37),
    'UG COMPRESSED AIR': (105, 110, 13),
    'UG MEDICAL AIR': (76, 153, 133),
    'UG MEDICAL VACUUM': (255, 0, 191),

    # duct
    'GEN EXH (-2 WG)': (0, 191, 255),
    'GEN EXH (-3 WG)': (0, 191, 255),
    'GEN EXH (-4 WG)': (0, 191, 255),
    'GEN EXH (-6 WG)': (0, 191, 255),
    'HEAT EXH (-1 WG)': (255, 0, 191),
    'HEAT EXH (-2 WG)': (255, 0, 191),
    'HEAT EXH (-3 WG)': (255, 0, 191),
    'OSA (+2 WG)': (255, 255, 0),
    'OSA (+3 WG)': (255, 255, 0),
    'RA (-1 WG)': (0, 255, 0),
    'RA (-2 WG)': (0, 255, 0),
    'RA (-3 WG)': (0, 255, 0),
    'RA (-4 WG)': (0, 255, 0),
    'RA (-6 WG)': (0, 255, 0),
    'SA (+2 WG)': (0, 255, 255),
    'SA (+3 WG)': (0, 255, 255),
    'SA (+4 WG)': (0, 255, 255),
    'SA (+6 WG)': (0, 255, 255),
    'SA (+10 WG)': (0, 255, 255),
    'RELF EXH (-2 WG)': (0, 94, 189),
    'RELF EXH (-3 WG)': (0, 94, 189),
    'RELF EXH (-4 WG)': (0, 94, 189),

# Add more system names and their corresponding RGB values here
}

# Define a dictionary to store custom filters (OrderedDict to preserve order)
custom_filters = OrderedDict()
# Adding items in the desired order
custom_filters["INSULATION"] = {
    "parameter_name": "Specification",
    "condition": "DoesNotContain",
    "value": "XYZ",
    "categories": [BuiltInCategory.OST_FabricationPipeworkInsulation, BuiltInCategory.OST_FabricationDuctworkInsulation, BuiltInCategory.OST_FabricationDuctworkLining],
    "color": (0, 0, 0)  # Black
}
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
custom_filters["BEAM HANGER"] = {
    "parameter_name": "FP_Beam Hanger",
    "condition": "Equals",
    "value": "Yes",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 255, 255)  # Cyan color for visibility
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
            elif condition == "DoesNotContain":
                # Create a Contains rule first
                contains_rule = ParameterFilterRuleFactory.CreateContainsRule(param_id, value, False)
                # Invert the Contains rule using the NOT operator
                rule = FilterInverseRule(contains_rule)
            elif condition == "Contains":
                # Create a Contains rule first
                contains_rule = ParameterFilterRuleFactory.CreateContainsRule(param_id, value, False)
            elif condition == "Equals":
                rule = ParameterFilterRuleFactory.CreateEqualsRule(param_id, value, False)
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
