import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')

from System.Collections.Generic import List
from Autodesk.Revit.DB import BuiltInCategory, Transaction, ElementId, ViewSchedule, FilteredElementCollector, ParameterElement, ScheduleFieldType, BuiltInParameter, ScheduleFilter, ScheduleFilterType, SpecTypeId, ScheduleSortGroupField, ScheduleSortOrder
from Autodesk.Revit.UI import TaskDialog
import System

# Define the active Revit application and document
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

# Function to check if a schedule with a specific name and category exists
def schedule_exists(schedule_name, category_id):
    schedules_collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    for schedule in schedules_collector:
        if schedule.Name == schedule_name and schedule.Definition.CategoryId == category_id:
            return True
    return False

# Function to check if all parameters for a schedule exist
def all_parameters_exist(field_names, parameters):
    for paramName, _ in field_names:
        if paramName in ["Type", "Family", "Elevation from Level", "Comments", "Level"]:
            continue  # These are built-in parameters, always available
        parameter = next((p for p in parameters if p.Name == paramName), None)
        if parameter is None:
            return False, paramName
    return True, None

# Define fields for the ROUND schedule
roundFieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("FP_Product Entry", "SIZE (OD of Pipe Including Insulation)"),
    ("Diameter", "SIZE (With Annular Space)"),
    ("Elevation from Level", "CL Elevation"),
    ("Type", "SLEEVE TYPE DR-WS=DROP WS=THRU"),  # Updated to use Type
    ("FP_Service Abbreviation", "SYSTEM ABBR."),
    ("FP_Service Name", "SERVICE NAME"),
    ("Level", "LEVEL")
]

# Define fields for the BLOCKOUT schedule
blockoutFieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("Width", "WIDTH"),
    ("Height", "HEIGHT"),
    ("Elevation from Level", "CL Elevation"),
    ("Family", "SLEEVE TYPE DR-WS=DROP WS=THRU"),
    ("Comments", "COMMENTS"),
    ("Level", "LEVEL")
]

# Check Revit version
revit_version = int(app.VersionNumber)
is_revit_2022_or_newer = revit_version >= 2022

file_name = doc.Title
categoryId = ElementId(BuiltInCategory.OST_PipeAccessory)

# Define schedule names and their respective family filters
schedules = [
    {"name": "WALL OPENING SCHEDULE", "fields": roundFieldNames, "filter": "WS", "category": categoryId},
    {"name": "WALL BLOCKOUT SCHEDULE", "fields": blockoutFieldNames, "filter": "BLOCKOUT", "category": categoryId}
]

# Function to add field by name
def add_field_by_name(definition, paramName, userColumnName, parameters):
    if paramName == "Type":
        paramId = ElementId(BuiltInParameter.ELEM_TYPE_PARAM)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
        field.ColumnHeading = userColumnName
    elif paramName == "Family":
        paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
        field.ColumnHeading = userColumnName
    elif paramName == "Elevation from Level":
        paramId = ElementId(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
        field.ColumnHeading = userColumnName
    elif paramName == "Comments":
        paramId = ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
        field.ColumnHeading = userColumnName
    elif paramName == "Level":
        paramId = ElementId(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
        field.ColumnHeading = userColumnName
    else:
        parameter = next((p for p in parameters if p.Name == paramName), None)
        if parameter is not None:
            paramId = parameter.Id
            field = definition.AddField(ScheduleFieldType.Instance, paramId)
            field.ColumnHeading = userColumnName
        else:
            TaskDialog.Show("Warning", "Parameter '{}' not found for schedule.".format(paramName))
    return field

# Start a new transaction
t = Transaction(doc, "Create Schedules")
t.Start()

# Get all parameters in the document
parameters = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

# Create schedules
for schedule_info in schedules:
    schedule_name = schedule_info["name"]
    fieldNames = schedule_info["fields"]
    family_filter = schedule_info["filter"]
    categoryId = schedule_info["category"]

    # Check if all parameters exist before creating the schedule
    params_exist, missing_param = all_parameters_exist(fieldNames, parameters)
    if not params_exist:
        TaskDialog.Show("Warning", "Cannot create '{}' because no blockouts were found.".format(schedule_name, missing_param))
        continue

    # Check if the schedule already exists
    if not schedule_exists(schedule_name, categoryId):
        # Create the schedule
        schedule = ViewSchedule.CreateSchedule(doc, categoryId)
        schedule.Name = schedule_name
        definition = schedule.Definition

        # Add fields to the schedule and store filter fields
        family_field = None
        ts_point_field = None
        type_field = None

        for paramName, userColumnName in fieldNames:
            field = add_field_by_name(definition, paramName, userColumnName, parameters)

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

            if paramName in ["Family", "Type"]:  # Updated to account for both Family and Type
                family_field = field

        definition.IsItemized = False

        if schedule_name == "WALL OPENING SCHEDULE" and ts_point_field is not None and type_field is not None:
            definition.AddSortGroupField(ScheduleSortGroupField(ts_point_field.FieldId, ScheduleSortOrder.Ascending))
            definition.AddSortGroupField(ScheduleSortGroupField(type_field.FieldId, ScheduleSortOrder.Ascending))

        # Add filters
        if is_revit_2022_or_newer:
            if schedule_name == "WALL OPENING SCHEDULE" and ts_point_field is not None and type_field is not None:
                definition.AddFilter(ScheduleFilter(ts_point_field.FieldId, ScheduleFilterType.HasValue))
                definition.AddFilter(ScheduleFilter(type_field.FieldId, ScheduleFilterType.NotEndsWith, "BLOCKOUT"))
            elif family_field is not None:
                schedule_filter = ScheduleFilter(
                    family_field.FieldId,
                    ScheduleFilterType.EndsWith,
                    family_filter
                )
                definition.AddFilter(schedule_filter)
    else:
        TaskDialog.Show("Schedule Exists", "'{}' already exists.".format(schedule_name))

t.Commit()