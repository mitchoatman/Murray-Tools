import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')

from Autodesk.Revit.DB import (
    BuiltInCategory, Transaction, ElementId, ViewSchedule,
    FilteredElementCollector, ParameterElement, ScheduleFieldType,
    BuiltInParameter, ScheduleFilter, ScheduleFilterType,
    FormatOptions, UnitTypeId
)
from Autodesk.Revit.UI import TaskDialog
import System

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

def schedule_exists(schedule_name, category_id):
    schedules_collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    for schedule in schedules_collector:
        if schedule.Name == schedule_name and schedule.Definition.CategoryId == category_id:
            return True
    return False

roundFieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("FP_Product Entry", "DIAMETER"),
    ("Diameter", "SIZE (With Annular Space)"),
    ("Elevation from Level", "CL Elevation"),
    ("Family", "SLEEVE TYPE DR-WS=DROP WS=THRU"),
    ("FP_Service Abbreviation", "SYSTEM ABBR."),
    ("FP_Service Name", "SERVICE NAME"),
    ("Comments", "COMMENTS")
]

rectFieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("Width", "WIDTH"),
    ("Height", "HEIGHT"),
    ("Elevation from Level", "CL Elevation"),
    ("Family", "SLEEVE TYPE DR-WS=DROP WS=THRU"),
    ("FP_Service Abbreviation", "SYSTEM ABBR."),
    ("FP_Service Name", "SERVICE NAME"),
    ("Comments", "COMMENTS")
]

revit_version = int(app.VersionNumber)
is_revit_2022_or_newer = revit_version >= 2022

categoryId = ElementId(BuiltInCategory.OST_DuctAccessory)

schedules = [
    {"name": "ROUND WALL SLEEVE SCHEDULE", "fields": roundFieldNames, "filter": "RDS", "category": categoryId},
    {"name": "RECTANGLE WALL SLEEVE SCHEDULE", "fields": rectFieldNames, "filter": "RWS", "category": categoryId}
]

def set_fractional_inches_1_8(field):
    fmt = field.GetFormatOptions()
    fmt.UseDefault = False
    fmt.SetUnitTypeId(UnitTypeId.FractionalInches)
    fmt.Accuracy = 0.125  # 1/8"
    field.SetFormatOptions(fmt)

def add_field_by_name(definition, paramName, userColumnName, parameters):
    field = None

    if paramName == "Family":
        paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
    elif paramName == "Elevation from Level":
        paramId = ElementId(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
    elif paramName == "Comments":
        paramId = ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
    else:
        parameter = next((p for p in parameters if p.Name == paramName), None)
        if parameter is not None:
            paramId = parameter.Id
            field = definition.AddField(ScheduleFieldType.Instance, paramId)
        else:
            TaskDialog.Show("Warning", "Parameter '{}' not found for schedule.".format(paramName))
            return None

    field.ColumnHeading = userColumnName

    if userColumnName in ["WIDTH", "HEIGHT", "SIZE (With Annular Space)"]:
        set_fractional_inches_1_8(field)

    return field

t = Transaction(doc, "Create Schedules")
t.Start()

parameters = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

for schedule_info in schedules:
    schedule_name = schedule_info["name"]
    fieldNames = schedule_info["fields"]
    family_filter = schedule_info["filter"]
    categoryId = schedule_info["category"]

    if not schedule_exists(schedule_name, categoryId):
        schedule = ViewSchedule.CreateSchedule(doc, categoryId)
        schedule.Name = schedule_name
        definition = schedule.Definition

        family_field = None
        for paramName, userColumnName in fieldNames:
            field = add_field_by_name(definition, paramName, userColumnName, parameters)
            if paramName == "Family" and field is not None:
                family_field = field

        if is_revit_2022_or_newer and family_field is not None:
            schedule_filter = ScheduleFilter(
                family_field.FieldId,
                ScheduleFilterType.EndsWith,
                family_filter
            )
            definition.AddFilter(schedule_filter)
    else:
        TaskDialog.Show("Schedule Exists", "'{}' already exists.".format(schedule_name))

t.Commit()