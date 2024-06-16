import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from random import randint
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.Attributes import *
from System.Collections.Generic import List
import System
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsValueString, get_parameter_value_by_name_AsInteger

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
categories.Add(ElementId(BuiltInCategory.OST_FabricationDuctwork))

# 3a. create rules and filters for each service name
fabrication_service_name_parameter = ElementId(BuiltInParameter.FABRICATION_SERVICE_NAME)

# Collect existing ParameterFilterElements
existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
existing_filter_names = {filter.Name for filter in existing_filters}

def random_color():
    r = randint(0, 230)
    g = randint(0, 230)
    b = randint(0, 230)
    return r, g, b

# Define a dictionary to map system names to their RGB values
system_colors = {
#waters
    'IRRIGATION WATER':(78, 131, 191),
    'DOMESTIC COLD WATER': (0, 63, 255),
    'DOMESTIC HOT WATER': (227, 34, 143),
    'HEATING WATER': (227, 148, 20),
    'UG DOMESTIC COLD WATER': (0, 63, 255),
    'CHILLED WATER': (133, 255, 190),
#waste and vent
    'SUMP PUMP DISCHARGE':(115, 117, 8),
    'TRAP PRIMER':(12, 38, 207),
    'OVERFLOW DRAIN':(179, 16, 146),
    'SANITARY VENT':(21, 237, 50),
    'SANITARY WASTE':(189, 0, 189),
    'CONDENSATE DRAIN': (26, 112, 66),
    'REFRIGERATION HOT GAS': (21, 130, 77),
    'CARBON DIOXIDE': (22, 107, 37),
    'STORM DRAIN': (157, 56, 224),
    'EMERGENCY DRAIN':(80, 4, 92),
    'GREASE WASTE':(131, 138, 58),
    'GREY WASTE':(138, 109, 58),
    'LAB WASTE':(55, 120, 81),
    'LAB VENT':(83, 184, 123),
    'UG CONDENSATE DRAIN': (26, 112, 66),
    'UG SANITARY VENT':(21, 237, 50),
    'UG SANTIARY WASTE':(153, 5, 143),
    'UG GREY WASTE':(138, 109, 58),
    'UG STORM DRAIN': (157, 56, 224),
    'UG LAB WASTE':(55, 120, 81),
    'UG LAB VENT':(83, 184, 123),
    'UG TRAP PRIMER':(12, 38, 207),
#gasses
    'LAB AIR':(43, 102, 67),
    'COMPRESSED AIR':(105, 110, 13),
    'GAS':(11, 97, 14),
    'NITROGEN':(41, 5, 66),
    'OXYGEN':(6, 59, 2),
    'UG CARBON DIOXIDE': (22, 107, 37),


    'UG OVERFLOW DRAIN':(179, 16, 146),
    # Add more system names and their corresponding RGB values here
}

with Transaction(doc, "Create and Apply Filters") as t:
    t.Start()

    for service_name in SNamelist_set:

        if service_name in system_colors:
            r, g, b = system_colors[service_name]
        else:
            r, g, b = random_color()

        if service_name not in existing_filter_names:
            rule = ParameterFilterRuleFactory.CreateEqualsRule(fabrication_service_name_parameter, service_name, False)
            filter = ElementParameterFilter(rule)

            # 5. create parameter filter 
            paramFilter = ParameterFilterElement.Create(doc, service_name, categories)
            paramFilter.SetElementFilter(filter)

            # 6. set graphic overrides
            overrides = OverrideGraphicSettings()
            overrides.SetProjectionLineColor(Color(r, g, b))
            #overrides.SetCutLineWeight(5)

            # 7. apply filter to view or view template and set visibility and overrides
            view_template_id = curview.ViewTemplateId
            if not view_template_id.Equals(ElementId.InvalidElementId):
                view_to_modify = doc.GetElement(view_template_id)
            else:
                view_to_modify = curview

            view_to_modify.AddFilter(paramFilter.Id)
            view_to_modify.SetFilterVisibility(paramFilter.Id, True)
            view_to_modify.SetFilterOverrides(paramFilter.Id, overrides)

    t.Commit()
