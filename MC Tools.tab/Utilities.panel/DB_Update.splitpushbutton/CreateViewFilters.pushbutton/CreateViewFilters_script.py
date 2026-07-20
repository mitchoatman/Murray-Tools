import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FabricationPart,
    FilteredElementCollector,
    ParameterFilterRuleFactory,
    Transaction,
    Color,
    LinePatternElement,
    BuiltInParameter,
    BuiltInCategory,
    ElementId,
    ParameterFilterElement,
    FilterInverseRule,
    ElementParameterFilter,
    OverrideGraphicSettings,
    FabricationConfiguration,
    View
)
from System.Collections.Generic import List
from collections import OrderedDict
from random import randint

from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)


def random_color():
    r = randint(0, 230)
    g = randint(0, 230)
    b = randint(0, 230)
    return r, g, b


def get_valid_rgb_from_overrides(overrides):
    try:
        col = overrides.ProjectionLineColor
        if col and col.IsValid:
            return (col.Red, col.Green, col.Blue)
    except:
        pass
    return None


def get_fallback_color(service_name, system_colors, service_color_dict):
    if service_name in system_colors:
        return system_colors[service_name]
    elif service_name in service_color_dict:
        return service_color_dict[service_name]
    else:
        return random_color()


def get_param_id_from_sample(doc, param_name):
    sample_element = FilteredElementCollector(doc).OfClass(FabricationPart).WhereElementIsNotElementType().FirstElement()
    if not sample_element:
        print 'No FabricationPart elements found to retrieve parameter ID for filter: ' + param_name
        return None

    for p in sample_element.Parameters:
        try:
            if p.Definition.Name == param_name:
                return p.Id
        except:
            pass
    return None


def build_rule(param_id, condition, value, revit_int):
    if revit_int < 2023:
        if condition == "EndsWith":
            return ParameterFilterRuleFactory.CreateEndsWithRule(param_id, value, False)
        elif condition == "DoesNotContain":
            contains_rule = ParameterFilterRuleFactory.CreateContainsRule(param_id, value, False)
            return FilterInverseRule(contains_rule)
        elif condition == "Equals":
            return ParameterFilterRuleFactory.CreateEqualsRule(param_id, value, False)
        elif condition == "Contains":
            return ParameterFilterRuleFactory.CreateContainsRule(param_id, value, False)
    else:
        if condition == "EndsWith":
            return ParameterFilterRuleFactory.CreateEndsWithRule(param_id, value)
        elif condition == "DoesNotContain":
            contains_rule = ParameterFilterRuleFactory.CreateContainsRule(param_id, value)
            return FilterInverseRule(contains_rule)
        elif condition == "Equals":
            return ParameterFilterRuleFactory.CreateEqualsRule(param_id, value)
        elif condition == "Contains":
            return ParameterFilterRuleFactory.CreateContainsRule(param_id, value)

    return None


# Check if user has fabrication part(s) selected
selected_ids = uidoc.Selection.GetElementIds()
selected_service_names = []

if selected_ids.Count > 0:
    for elem_id in selected_ids:
        selected_element = doc.GetElement(elem_id)
        if isinstance(selected_element, FabricationPart):
            service_name = selected_element.get_Parameter(BuiltInParameter.FABRICATION_SERVICE_NAME).AsString()
            if service_name and service_name not in selected_service_names:
                selected_service_names.append(service_name)

    if selected_service_names:
        SNamelist = selected_service_names
        SNamelist_set = set(SNamelist)
    else:
        SNamelist = []
        Config = FabricationConfiguration.GetFabricationConfiguration(doc)
        if Config:
            LoadedServices = Config.GetAllUsedServices()
            SNamelist = [service.Name.split(':')[1].strip() for service in LoadedServices if service.Name and ':' in service.Name]
        else:
            print 'No Fabrication Configuration found in the project'
            raise Exception('No Fabrication Configuration found in the project')
        SNamelist_set = set(SNamelist)
else:
    SNamelist = []
    Config = FabricationConfiguration.GetFabricationConfiguration(doc)
    if Config:
        LoadedServices = Config.GetAllUsedServices()
        SNamelist = [service.Name.split(':')[1].strip() for service in LoadedServices if service.Name and ':' in service.Name]
    else:
        print 'No Fabrication Configuration found in the project'
        raise Exception('No Fabrication Configuration found in the project')
    SNamelist_set = set(SNamelist)

# Create list of categories that will be used for service filters
categories = List[ElementId]()
categories.Add(ElementId(BuiltInCategory.OST_FabricationHangers))
categories.Add(ElementId(BuiltInCategory.OST_FabricationPipework))
categories.Add(ElementId(BuiltInCategory.OST_FabricationDuctwork))

# Fabrication service name parameter
fabrication_service_name_parameter = ElementId(BuiltInParameter.FABRICATION_SERVICE_NAME)

# Collect existing ParameterFilterElements
existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
existing_filter_names = {f.Name for f in existing_filters}
existing_filter_dict = {f.Name: f.Id for f in existing_filters}

# Look for Dash line pattern
line_patterns = FilteredElementCollector(doc).OfClass(LinePatternElement)
dashed_pattern_id = None
for pattern in line_patterns:
    if pattern.Name == "Dash":
        dashed_pattern_id = pattern.Id
        break

if dashed_pattern_id is None:
    raise ValueError("Dashed line pattern not found in the document.")

# Define system colors
system_colors = {
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

    'ACID WASTE': (189, 189, 126),
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
    'UG SANITARY WASTE': (204, 0, 204),
    'UG GREY WASTE': (204, 153, 102),
    'UG STORM DRAIN': (255, 127, 191),
    'UG LAB WASTE': (55, 120, 81),
    'UG LAB VENT': (83, 184, 123),
    'UG TRAP PRIMER': (12, 38, 207),
    'UG GREASE WASTE': (189, 189, 126),
    'FUEL OIL SUPPLY': (175, 175, 0),

    'CARBON DIOXIDE': (22, 107, 37),
    'LAB AIR': (76, 153, 133),
    'LAB VACUUM': (255, 0, 191),
    'COMPRESSED AIR': (105, 110, 13),
    'MEDICAL AIR': (76, 153, 133),
    'MEDICAL VACUUM': (255, 0, 191),
    'MEDICAL NITROGEN': (41, 5, 66),
    'MEDICAL OXYGEN': (6, 59, 2),
    'NITROGEN': (41, 5, 66),
    'OXYGEN': (6, 59, 2),
    'GAS': (0, 129, 0),
    'GAS VENT': (170, 255, 191),
    'GAS - LOW PRESSURE': (0, 129, 0),
    'UG NITROGEN': (41, 5, 66),
    'UG OXYGEN': (6, 59, 2),
    'UG CARBON DIOXIDE': (22, 107, 37),
    'UG COMPRESSED AIR': (105, 110, 13),
    'UG MEDICAL AIR': (76, 153, 133),
    'UG MEDICAL VACUUM': (255, 0, 191),

    'GEN EXH (-2 WG)': (0, 191, 255),
    'GEN EXH (-3 WG)': (0, 191, 255),
    'GEN EXH (-4 WG)': (0, 191, 255),
    'GEN EXH (-6 WG)': (0, 191, 255),
    'HEAT EXH (-1 WG)': (255, 0, 191),
    'HEAT EXH (-2 WG)': (255, 0, 191),
    'HEAT EXH (-3 WG)': (255, 0, 191),
    'OSA (+2 WG)': (255, 128, 64),
    'OSA (+3 WG)': (255, 128, 64),
    'OSA (+4 WG)': (255, 128, 64),
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
}

# Custom filters
custom_filters = OrderedDict()
custom_filters["FP_INSULATION"] = {
    "parameter_name": "Specification",
    "condition": "DoesNotContain",
    "value": "XYZ",
    "categories": [
        BuiltInCategory.OST_FabricationPipeworkInsulation,
        BuiltInCategory.OST_FabricationDuctworkInsulation,
        BuiltInCategory.OST_FabricationDuctworkLining
    ],
    "color": (0, 0, 0)
}
custom_filters["REVIEW"] = {
    "parameter_name": "Comments",
    "condition": "Contains",
    "value": "REVIEW",
    "categories": [
        BuiltInCategory.OST_FabricationPipework,
        BuiltInCategory.OST_FabricationHangers,
        BuiltInCategory.OST_FabricationDuctwork
    ],
    "color": (255, 165, 0)
}
custom_filters["PLACEHOLDER"] = {
    "parameter_name": "Comments",
    "condition": "Contains",
    "value": "PLACEHOLDER",
    "categories": [
        BuiltInCategory.OST_FabricationPipework,
        BuiltInCategory.OST_FabricationHangers,
        BuiltInCategory.OST_FabricationDuctwork
    ],
    "color": (255, 0, 0)
}
custom_filters["SPOOL 1"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "1",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (255, 0, 0)
}
custom_filters["SPOOL 2"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "2",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (79, 0, 59)
}
custom_filters["SPOOL 3"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "3",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 0, 255)
}
custom_filters["SPOOL 4"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "4",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 255, 0)
}
custom_filters["SPOOL 5"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "5",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (255, 0, 0)
}
custom_filters["SPOOL 6"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "6",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (79, 0, 59)
}
custom_filters["SPOOL 7"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "7",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 0, 255)
}
custom_filters["SPOOL 8"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "8",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 255, 0)
}
custom_filters["SPOOL 9"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "9",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (255, 0, 0)
}
custom_filters["SPOOL 0"] = {
    "parameter_name": "STRATUS Assembly",
    "condition": "EndsWith",
    "value": "0",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (79, 0, 59)
}
custom_filters["BEAM HANGER"] = {
    "parameter_name": "FP_Beam Hanger",
    "condition": "Equals",
    "value": "Yes",
    "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
    "color": (0, 255, 255)
}

# Determine whether to modify current view or its template
view_template_id = curview.ViewTemplateId
if not view_template_id.Equals(ElementId.InvalidElementId):
    view_to_modify = doc.GetElement(view_template_id)
else:
    view_to_modify = curview

# Get filters already applied to target view/template
applied_filters = {doc.GetElement(f_id).Name: f_id for f_id in view_to_modify.GetFilters()}

# Pre-collect colors for service filters from other views/templates
service_color_dict = {}
views_collector = FilteredElementCollector(doc).OfClass(View)

for view_elem in views_collector:
    if view_elem.Id.Equals(view_to_modify.Id):
        continue

    try:
        applied_filter_ids = view_elem.GetFilters()
        for f_id in applied_filter_ids:
            f_elem = doc.GetElement(f_id)
            if f_elem and f_elem.Name in SNamelist_set:
                try:
                    ovr = view_elem.GetFilterOverrides(f_id)
                    rgb = get_valid_rgb_from_overrides(ovr)
                    if rgb and f_elem.Name not in service_color_dict:
                        service_color_dict[f_elem.Name] = rgb
                except:
                    pass
    except:
        pass

with Transaction(doc, "Create and Apply Filters") as t:
    t.Start()

    # Create/apply custom filters
    for filter_name, filter_props in custom_filters.items():
        param_name = filter_props["parameter_name"]
        condition = filter_props["condition"]
        value = filter_props["value"]
        filter_categories = List[ElementId]([ElementId(cat) for cat in filter_props["categories"]])
        color = Color(*filter_props["color"])

        if filter_name in existing_filter_names:
            filter_id = existing_filter_dict[filter_name]
        else:
            param_id = get_param_id_from_sample(doc, param_name)
            if not param_id:
                print 'Parameter ' + param_name + ' not found for filter: ' + filter_name
                continue

            rule = build_rule(param_id, condition, value, RevitINT)
            if not rule:
                print 'Condition not supported for filter: ' + filter_name
                continue

            filter_element = ElementParameterFilter(rule)
            filter_elem = ParameterFilterElement.Create(doc, filter_name, filter_categories)
            filter_elem.SetElementFilter(filter_element)
            filter_id = filter_elem.Id

        overrides = OverrideGraphicSettings()
        overrides.SetProjectionLineColor(color)

        if filter_name == "FP_INSULATION":
            overrides.SetSurfaceTransparency(100)
            overrides.SetProjectionLinePatternId(dashed_pattern_id)
            overrides.SetHalftone(True)

        if filter_name not in applied_filters:
            view_to_modify.AddFilter(filter_id)
            view_to_modify.SetFilterVisibility(filter_id, True)
            if filter_name.startswith("SPOOL "):
                view_to_modify.SetIsFilterEnabled(filter_id, False)

        view_to_modify.SetFilterOverrides(filter_id, overrides)

    # Create/apply service filters
    for service_name in SNamelist_set:
        if not service_name:
            print 'Skipping empty or None service name'
            continue

        if service_name in existing_filter_names:
            paramFilterId = existing_filter_dict[service_name]

            if service_name in applied_filters:
                existing_overrides = view_to_modify.GetFilterOverrides(paramFilterId)
                rgb = get_valid_rgb_from_overrides(existing_overrides)

                if rgb:
                    r, g, b = rgb
                else:
                    r, g, b = get_fallback_color(service_name, system_colors, service_color_dict)
            else:
                r, g, b = get_fallback_color(service_name, system_colors, service_color_dict)

        else:
            r, g, b = get_fallback_color(service_name, system_colors, service_color_dict)

            try:
                if RevitINT < 2023:
                    rule = ParameterFilterRuleFactory.CreateEqualsRule(fabrication_service_name_parameter, service_name, False)
                else:
                    rule = ParameterFilterRuleFactory.CreateEqualsRule(fabrication_service_name_parameter, service_name)

                filter_rule = ElementParameterFilter(rule)
                paramFilter = ParameterFilterElement.Create(doc, service_name, categories)
                paramFilter.SetElementFilter(filter_rule)
                paramFilterId = paramFilter.Id
            except Exception as e:
                print 'Failed to create filter for service: ' + service_name + '. Error: ' + str(e)
                continue

        overrides = OverrideGraphicSettings()
        overrides.SetProjectionLineColor(Color(r, g, b))

        if service_name not in applied_filters:
            try:
                view_to_modify.AddFilter(paramFilterId)
                view_to_modify.SetFilterVisibility(paramFilterId, True)
            except Exception as e:
                print 'Failed to apply filter for service: ' + service_name + ' to view. Error: ' + str(e)
                continue

        try:
            view_to_modify.SetFilterOverrides(paramFilterId, overrides)
        except Exception as e:
            print 'Failed to set overrides for service: ' + service_name + '. Error: ' + str(e)

    try:
        t.Commit()
    except Exception as e:
        print 'Transaction failed to commit. Error: ' + str(e)
        t.RollBack()