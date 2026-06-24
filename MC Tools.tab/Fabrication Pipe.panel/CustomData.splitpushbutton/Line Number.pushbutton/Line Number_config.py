import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FabricationPart,
    FilteredElementCollector,
    ParameterFilterRuleFactory,
    Transaction,
    Color,
    BuiltInCategory,
    ElementId,
    ParameterFilterElement,
    ElementParameterFilter,
    OverrideGraphicSettings,
    View
)
from System.Collections.Generic import List
from random import randint

from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

PARAM_NAME = "FP_Line Number"
FILTER_PREFIX = "LINE - "   # change to "" if you want filter names to be only the line number


def natural_key(s):
    import re
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'([0-9]+)', s)]


def random_color():
    r = randint(0, 230)
    g = randint(0, 230)
    b = randint(0, 230)
    return r, g, b


def get_target_view_or_template():
    view_template_id = curview.ViewTemplateId
    if not view_template_id.Equals(ElementId.InvalidElementId):
        return doc.GetElement(view_template_id)
    return curview


def get_line_number_param_id():
    sample_element = FilteredElementCollector(doc).OfClass(FabricationPart).WhereElementIsNotElementType().FirstElement()
    if not sample_element:
        return None

    for p in sample_element.Parameters:
        try:
            if p.Definition.Name == PARAM_NAME:
                return p.Id
        except:
            pass
    return None


def get_all_line_numbers_in_project():
    values = set()
    collector = FilteredElementCollector(doc).OfClass(FabricationPart).WhereElementIsNotElementType()

    for elem in collector:
        try:
            param = elem.LookupParameter(PARAM_NAME)
            if param and param.HasValue:
                val = param.AsString()
                if val and val.strip():
                    values.add(val.strip())
        except:
            pass

    return sorted(values, key=natural_key)


# categories for the filters
categories = List[ElementId]()
categories.Add(ElementId(BuiltInCategory.OST_FabricationHangers))
categories.Add(ElementId(BuiltInCategory.OST_FabricationPipework))
categories.Add(ElementId(BuiltInCategory.OST_FabricationDuctwork))

view_to_modify = get_target_view_or_template()

# existing filters in model
existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
existing_filter_names = {f.Name for f in existing_filters}
existing_filter_dict = {f.Name: f.Id for f in existing_filters}

# filters already applied to target
applied_filters = {}
try:
    applied_filters = {doc.GetElement(fid).Name: fid for fid in view_to_modify.GetFilters()}
except:
    applied_filters = {}

# collect existing line filter colors from other views/templates
line_filter_color_dict = {}
views_collector = FilteredElementCollector(doc).OfClass(View)

for view_elem in views_collector:
    if view_elem.Id.Equals(view_to_modify.Id):
        continue

    try:
        applied_filter_ids = view_elem.GetFilters()
        for f_id in applied_filter_ids:
            f_elem = doc.GetElement(f_id)
            if f_elem and f_elem.Name.startswith(FILTER_PREFIX):
                if f_elem.Name not in line_filter_color_dict:
                    try:
                        ovr = view_elem.GetFilterOverrides(f_id)
                        col = ovr.ProjectionLineColor
                        line_filter_color_dict[f_elem.Name] = (col.Red, col.Green, col.Blue)
                    except:
                        pass
    except:
        pass

line_numbers = get_all_line_numbers_in_project()
if not line_numbers:
    raise Exception("No '{}' values found on Fabrication Parts in the project.".format(PARAM_NAME))

param_id = get_line_number_param_id()
if not param_id:
    raise Exception("Could not find parameter id for '{}'.".format(PARAM_NAME))

created_count = 0
applied_count = 0
kept_existing_count = 0

with Transaction(doc, "Create and Apply Line Number Filters") as t:
    t.Start()

    for line_number in line_numbers:
        filter_name = FILTER_PREFIX + line_number
        filter_id = None
        keep_existing_overrides = False
        rgb = None

        # filter exists already
        if filter_name in existing_filter_names:
            filter_id = existing_filter_dict[filter_name]

            # already applied to target: keep current color/overrides
            if filter_name in applied_filters:
                keep_existing_overrides = True
                kept_existing_count += 1
            else:
                if filter_name in line_filter_color_dict:
                    rgb = line_filter_color_dict[filter_name]
                else:
                    rgb = random_color()

        # create new filter
        else:
            try:
                if RevitINT < 2023:
                    rule = ParameterFilterRuleFactory.CreateEqualsRule(param_id, line_number, False)
                else:
                    rule = ParameterFilterRuleFactory.CreateEqualsRule(param_id, line_number)

                elem_filter = ElementParameterFilter(rule)
                param_filter = ParameterFilterElement.Create(doc, filter_name, categories)
                param_filter.SetElementFilter(elem_filter)
                filter_id = param_filter.Id
                created_count += 1

                if filter_name in line_filter_color_dict:
                    rgb = line_filter_color_dict[filter_name]
                else:
                    rgb = random_color()

            except Exception as e:
                print 'Failed to create filter for line number: {}. Error: {}'.format(line_number, str(e))
                continue

        # apply to template/view if not already applied
        if filter_name not in applied_filters:
            try:
                view_to_modify.AddFilter(filter_id)
                view_to_modify.SetFilterVisibility(filter_id, True)
                applied_count += 1
            except Exception as e:
                print 'Failed to apply filter: {}. Error: {}'.format(filter_name, str(e))
                continue

        # only set overrides if target does not already have them
        if not keep_existing_overrides and rgb is not None:
            try:
                overrides = OverrideGraphicSettings()
                overrides.SetProjectionLineColor(Color(rgb[0], rgb[1], rgb[2]))
                view_to_modify.SetFilterOverrides(filter_id, overrides)
            except Exception as e:
                print 'Failed to set overrides for filter: {}. Error: {}'.format(filter_name, str(e))

    try:
        t.Commit()
        print 'Line number filters complete.'
        print 'Created: {}'.format(created_count)
        print 'Applied to target: {}'.format(applied_count)
        print 'Kept existing target overrides: {}'.format(kept_existing_count)
    except Exception as e:
        print 'Transaction failed to commit. Error: ' + str(e)
        t.RollBack()