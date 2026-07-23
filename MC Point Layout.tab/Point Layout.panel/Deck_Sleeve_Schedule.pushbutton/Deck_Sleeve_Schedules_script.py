import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitAPIUI')

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
    ScheduleSortGroupField,
    ScheduleSortOrder
)
from Autodesk.Revit.UI import TaskDialog
import System

# Active Revit document/UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

if not file_name:
    file_name = doc.Title

fieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("Diameter", "SIZE"),
    ("Length", "LENGTH"),
    ("Family", "NAME")
]

categoryId = ElementId(BuiltInCategory.OST_PipeAccessory)
schedule_name = "DECK SLEEVE SCHEDULE"

revit_version = int(app.VersionNumber)
is_revit_2022_or_newer = revit_version >= 2022
is_revit_2023_or_newer = revit_version >= 2023


def get_id_value(eid):
    try:
        return eid.Value        # Revit 2024+
    except:
        return eid.IntegerValue # older Revit


def get_existing_schedule(schedule_name, category_id):
    schedules = FilteredElementCollector(doc).OfClass(ViewSchedule)
    for schedule in schedules:
        try:
            if schedule.Name == schedule_name and schedule.Definition.CategoryId == category_id:
                return schedule
        except:
            pass
    return None


def open_view(view):
    if not view or not isinstance(view, View):
        TaskDialog.Show("Open View", "View is not valid.")
        return

    if view.IsTemplate:
        TaskDialog.Show("Open View", "Cannot open a view template.")
        return

    if isinstance(view, ViewSchedule):
        if view.IsInternalKeynoteSchedule or view.IsTitleblockRevisionSchedule:
            TaskDialog.Show("Open View", "Cannot open this internal schedule view.")
            return

    try:
        uidoc.ActiveView = view
    except Exception as ex:
        TaskDialog.Show("Open View", "Could not open view '{}'\n{}".format(view.Name, ex))


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

    if paramName == "Family":
        paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
        existing = get_schedule_field_by_parameter_id(definition, paramId)
        if existing:
            existing.ColumnHeading = userColumnName
            return existing

        field = definition.AddField(ScheduleFieldType.Instance, paramId)
        field.ColumnHeading = userColumnName

    elif paramName == "Elevation from Level":
        paramId = ElementId(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
        existing = get_schedule_field_by_parameter_id(definition, paramId)
        if existing:
            existing.ColumnHeading = userColumnName
            return existing

        field = definition.AddField(ScheduleFieldType.Instance, paramId)
        field.ColumnHeading = userColumnName

    else:
        parameter = get_parameter_element_by_name(paramName, parameters)
        if parameter is not None:
            paramId = parameter.Id
            existing = get_schedule_field_by_parameter_id(definition, paramId)
            if existing:
                existing.ColumnHeading = userColumnName
                return existing

            field = definition.AddField(ScheduleFieldType.Instance, paramId)
            field.ColumnHeading = userColumnName

    return field


def configure_schedule(schedule, parameters):
    definition = schedule.Definition

    # Sort by TS_Point_Number
    ts_param = get_parameter_element_by_name("TS_Point_Number", parameters)
    if ts_param is not None:
        ts_field = get_schedule_field_by_parameter_id(definition, ts_param.Id)
        if ts_field is not None:
            definition.ClearSortGroupFields()
            sort_field = ScheduleSortGroupField(ts_field.FieldId, ScheduleSortOrder.Ascending)
            definition.AddSortGroupField(sort_field)

    # Revit 2023+ : Filter by Sheet
    if is_revit_2023_or_newer:
        try:
            if definition.IsValidCategoryForFilterBySheet():
                definition.IsFilteredBySheet = True
        except:
            try:
                definition.IsFilteredBySheet = True
            except:
                pass


# Find the ElementId of the "Round Floor Sleeve" family type
round_floor_sleeve_type_id = None
if is_revit_2022_or_newer:
    pipe_accessories = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeAccessory)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for elem in pipe_accessories:
        fam_param = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
        if fam_param and fam_param.AsValueString() == "Round Floor Sleeve":
            round_floor_sleeve_type_id = elem.GetTypeId()
            break

parameters = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()
existing_schedule = get_existing_schedule(schedule_name, categoryId)

if existing_schedule:
    t = Transaction(doc, "Update Deck Sleeve Schedule")
    t.Start()

    configure_schedule(existing_schedule, parameters)

    t.Commit()
    open_view(existing_schedule)

else:
    t = Transaction(doc, "Create Schedule")
    t.Start()

    schedule = ViewSchedule.CreateSchedule(doc, categoryId)
    schedule.Name = schedule_name
    definition = schedule.Definition

    for paramName, userColumnName in fieldNames:
        add_field_by_name(definition, paramName, userColumnName, parameters)

    # Add filter for "Round Floor Sleeve" if Revit 2022 or newer
    if is_revit_2022_or_newer and round_floor_sleeve_type_id is not None:
        type_param_id = ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
        type_field = get_schedule_field_by_parameter_id(definition, type_param_id)

        if type_field is None:
            type_field = definition.AddField(ScheduleFieldType.Instance, type_param_id)
            type_field.ColumnHeading = "Family and Type"
            type_field.IsHidden = True

        schedule_filter = ScheduleFilter(
            type_field.FieldId,
            ScheduleFilterType.Equal,
            round_floor_sleeve_type_id
        )
        definition.AddFilter(schedule_filter)

    configure_schedule(schedule, parameters)

    t.Commit()
    open_view(schedule)