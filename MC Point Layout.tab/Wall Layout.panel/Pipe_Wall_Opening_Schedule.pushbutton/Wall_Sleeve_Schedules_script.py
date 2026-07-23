import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Transaction,
    ElementId,
    View,
    ViewSchedule,
    FilteredElementCollector,
    ParameterElement,
    ScheduleFieldType,
    BuiltInParameter,
    ScheduleFilter,
    ScheduleFilterType,
    SpecTypeId,
    ScheduleSortGroupField,
    ScheduleSortOrder
)
from Autodesk.Revit.UI import TaskDialog
import System

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)
if not file_name:
    file_name = doc.Title

revit_version = int(app.VersionNumber)
is_revit_2022_or_newer = revit_version >= 2022
is_revit_2023_or_newer = revit_version >= 2023

categoryId = ElementId(BuiltInCategory.OST_PipeAccessory)

roundFieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("FP_Product Entry", "SIZE (OD of Pipe Including Insulation)"),
    ("Diameter", "SIZE (With Annular Space)"),
    ("Elevation from Level", "CL Elevation"),
    ("Type", "SLEEVE TYPE DR-WS=DROP WS=THRU"),
    ("FP_Service Abbreviation", "SYSTEM ABBR."),
    ("FP_Service Name", "SERVICE NAME"),
    ("Level", "LEVEL")
]

blockoutFieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("Width", "WIDTH"),
    ("Height", "HEIGHT"),
    ("Elevation from Level", "CL Elevation"),
    ("Family", "SLEEVE TYPE DR-WS=DROP WS=THRU"),
    ("Comments", "COMMENTS"),
    ("Level", "LEVEL")
]

schedules = [
    {"name": "WALL OPENING SCHEDULE", "fields": roundFieldNames, "filter": "WS", "category": categoryId},
    {"name": "WALL BLOCKOUT SCHEDULE", "fields": blockoutFieldNames, "filter": "BLOCKOUT", "category": categoryId}
]

def get_id_value(eid):
    try:
        return eid.Value
    except:
        return eid.IntegerValue

def get_existing_schedule(schedule_name, category_id):
    schedules_collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    for schedule in schedules_collector:
        try:
            if schedule.Name == schedule_name and schedule.Definition.CategoryId == category_id:
                return schedule
        except:
            pass
    return None

def open_view(view):
    if not view or not isinstance(view, View):
        return

    if view.IsTemplate:
        return

    if isinstance(view, ViewSchedule):
        if view.IsInternalKeynoteSchedule or view.IsTitleblockRevisionSchedule:
            return

    try:
        uidoc.ActiveView = view
    except Exception as ex:
        TaskDialog.Show("Open View", "Could not open view '{}'\n{}".format(view.Name, ex))

def all_parameters_exist(field_names, parameters):
    for paramName, _ in field_names:
        if paramName in ["Type", "Family", "Elevation from Level", "Comments", "Level"]:
            continue
        parameter = next((p for p in parameters if p.Name == paramName), None)
        if parameter is None:
            return False, paramName
    return True, None

def get_parameter_element_by_name(param_name, parameters):
    return next((p for p in parameters if p.Name == param_name), None)

def get_schedule_field_by_parameter_id(definition, param_id):
    for field_id in definition.GetFieldOrder():
        field = definition.GetField(field_id)
        try:
            if get_id_value(field.ParameterId) == get_id_value(param_id):
                return field
        except:
            pass
    return None

def add_field_by_name(definition, paramName, userColumnName, parameters):
    field = None

    if paramName == "Type":
        paramId = ElementId(BuiltInParameter.ELEM_TYPE_PARAM)
    elif paramName == "Family":
        paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
    elif paramName == "Elevation from Level":
        paramId = ElementId(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
    elif paramName == "Comments":
        paramId = ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    elif paramName == "Level":
        paramId = ElementId(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
    else:
        parameter = next((p for p in parameters if p.Name == paramName), None)
        if parameter is None:
            TaskDialog.Show("Warning", "Parameter '{}' not found for schedule.".format(paramName))
            return None
        paramId = parameter.Id

    existing = get_schedule_field_by_parameter_id(definition, paramId)
    if existing:
        existing.ColumnHeading = userColumnName
        return existing

    field = definition.AddField(ScheduleFieldType.Instance, paramId)
    field.ColumnHeading = userColumnName
    return field

def configure_schedule(schedule, fieldNames, family_filter, parameters):
    definition = schedule.Definition
    definition.IsItemized = False

    family_field = None
    ts_point_field = None
    type_field = None

    for paramName, userColumnName in fieldNames:
        field = add_field_by_name(definition, paramName, userColumnName, parameters)
        if field is None:
            continue

        if paramName in ["Diameter", "Width", "Height"]:
            fmt = doc.GetUnits().GetFormatOptions(SpecTypeId.Length)
            fmt.UseDefault = False
            if fmt.CanSuppressLeadingZeros():
                fmt.SuppressLeadingZeros = True
            field.SetFormatOptions(fmt)

        if paramName == "TS_Point_Number":
            ts_point_field = field

        if paramName == "Type":
            type_field = field

        if paramName in ["Family", "Type"]:
            family_field = field

    # Sort by TS_Point_Number
    if ts_point_field is not None:
        try:
            definition.ClearSortGroupFields()
        except:
            pass
        definition.AddSortGroupField(
            ScheduleSortGroupField(ts_point_field.FieldId, ScheduleSortOrder.Ascending)
        )

    # Filter by Sheet for Revit 2023+
    if is_revit_2023_or_newer:
        try:
            if definition.IsValidCategoryForFilterBySheet():
                definition.IsFilteredBySheet = True
        except:
            try:
                definition.IsFilteredBySheet = True
            except:
                pass

    # Rebuild filters
    try:
        definition.ClearFilters()
    except:
        pass

    if is_revit_2022_or_newer:
        if schedule.Name == "WALL OPENING SCHEDULE" and ts_point_field is not None and type_field is not None:
            definition.AddFilter(ScheduleFilter(ts_point_field.FieldId, ScheduleFilterType.HasValue))
            definition.AddFilter(ScheduleFilter(type_field.FieldId, ScheduleFilterType.NotEndsWith, "BLOCKOUT"))
        elif family_field is not None:
            definition.AddFilter(
                ScheduleFilter(
                    family_field.FieldId,
                    ScheduleFilterType.EndsWith,
                    family_filter
                )
            )

parameters = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

t = Transaction(doc, "Create / Update Schedules")
t.Start()

last_schedule = None

for schedule_info in schedules:
    schedule_name = schedule_info["name"]
    fieldNames = schedule_info["fields"]
    family_filter = schedule_info["filter"]
    categoryId = schedule_info["category"]

    params_exist, missing_param = all_parameters_exist(fieldNames, parameters)
    if not params_exist:
        TaskDialog.Show("Warning", "Cannot create/update '{}' because parameter '{}' was not found.".format(schedule_name, missing_param))
        continue

    existing_schedule = get_existing_schedule(schedule_name, categoryId)

    if existing_schedule:
        configure_schedule(existing_schedule, fieldNames, family_filter, parameters)
        last_schedule = existing_schedule
    else:
        schedule = ViewSchedule.CreateSchedule(doc, categoryId)
        schedule.Name = schedule_name
        configure_schedule(schedule, fieldNames, family_filter, parameters)
        last_schedule = schedule

t.Commit()

if last_schedule:
    open_view(last_schedule)