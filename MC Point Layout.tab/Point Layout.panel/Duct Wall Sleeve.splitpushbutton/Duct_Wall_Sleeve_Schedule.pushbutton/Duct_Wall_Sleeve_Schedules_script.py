import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')

from System.Collections.Generic import List
from Autodesk.Revit.DB import BuiltInCategory, Transaction, ElementId, ViewSchedule, FilteredElementCollector, ParameterElement, ScheduleFieldType, BuiltInParameter, ScheduleFilter, ScheduleFilterType
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

# Define fields for the ROUND schedule
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

# Define fields for the RECTANGLE schedule
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

# Check Revit version
revit_version = int(app.VersionNumber)
is_revit_2022_or_newer = revit_version >= 2022

file_name = doc.Title
categoryId = ElementId(BuiltInCategory.OST_DuctAccessory)

# Define schedule names and their respective family filters
schedules = [
    {"name": "ROUND WALL SLEEVE SCHEDULE", "fields": roundFieldNames, "filter": "RDS", "category": categoryId},
    {"name": "RECTANGLE WALL SLEEVE SCHEDULE", "fields": rectFieldNames, "filter": "RWS", "category": categoryId}
]

# Function to add field by name
def add_field_by_name(definition, paramName, userColumnName, parameters):
    if paramName == "Family":
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

    # Check if the schedule already exists
    if not schedule_exists(schedule_name, categoryId):
        # Create the schedule
        schedule = ViewSchedule.CreateSchedule(doc, categoryId)
        schedule.Name = schedule_name
        definition = schedule.Definition

        # Add fields to the schedule and store the family field
        family_field = None
        for paramName, userColumnName in fieldNames:
            field = add_field_by_name(definition, paramName, userColumnName, parameters)
            if paramName == "Family":
                family_field = field

        # Add filter for family names ending with specific suffix if Revit 2022 or newer
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