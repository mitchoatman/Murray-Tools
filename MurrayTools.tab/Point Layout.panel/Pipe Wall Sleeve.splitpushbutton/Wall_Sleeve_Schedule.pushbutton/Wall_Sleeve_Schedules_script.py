import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')

from System.Collections.Generic import List
from Autodesk.Revit.DB import BuiltInCategory, Transaction, ElementId, ViewSchedule, FilteredElementCollector, ParameterElement, ScheduleFieldType, BuiltInParameter, ScheduleFilter, ScheduleFilterType
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

# Define the fields for the schedule and their desired order
fieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("FP_Product Entry", "SIZE (OD of Pipe Including Insulation)"),
    ("Diameter", "SIZE (With Annular Space)"),
    ("Elevation from Level", "CL Elevation"),
    ("Family", "SLEEVE TYPE DR-WS=DROP WS=THRU"),
    ("FP_Service Abbreviation", "SYSTEM ABBR."),
    ("FP_Service Name", "SERVICE NAME")
]

# Check Revit version
revit_version = int(app.VersionNumber)
is_revit_2022_or_newer = revit_version >= 2022

# Determine category and schedule name based on shift click
if __shiftclick__:
    file_name = doc.Title
    categoryId = ElementId(BuiltInCategory.OST_DuctAccessory)
    schedule_name = "WALL SLEEVE SCHEDULE"
else:
    categoryId = ElementId(BuiltInCategory.OST_PipeAccessory)
    schedule_name = "WALL SLEEVE SCHEDULE"

# Check if the schedule already exists
if not schedule_exists(schedule_name, categoryId):
    # Start a new transaction
    t = Transaction(doc, "Create Schedule")
    t.Start()

    # Create the schedule
    schedule = ViewSchedule.CreateSchedule(doc, categoryId)

    # Set the schedule name
    schedule.Name = schedule_name

    # Get the ScheduleDefinition from the ViewSchedule
    definition = schedule.Definition

    # Get all parameters in the document
    parameters = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

    # Function to add field by name
    def add_field_by_name(definition, paramName, userColumnName, parameters):
        if paramName == "Family":
            # Handle the Family parameter separately using the built-in parameter
            paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
            field = definition.AddField(ScheduleFieldType.Instance, paramId)
            field.ColumnHeading = userColumnName
        elif paramName == "Elevation from Level":
            paramId = ElementId(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
            field = definition.AddField(ScheduleFieldType.Instance, paramId)
            field.ColumnHeading = userColumnName
        else:
            # Find the parameter with the matching name
            parameter = next((p for p in parameters if p.Name == paramName), None)
            # Check if the parameter was found
            if parameter is not None:
                # Get the id of the parameter
                paramId = parameter.Id
                # Create a new ScheduleField from the parameter id
                field = definition.AddField(ScheduleFieldType.Instance, paramId)
                # Set the field column header
                field.ColumnHeading = userColumnName
        return field

    # Add fields to the schedule and store the family field
    family_field = None
    for paramName, userColumnName in fieldNames:
        field = add_field_by_name(definition, paramName, userColumnName, parameters)
        if paramName == "Family":
            family_field = field

    # Add filter for family names ending with "WS" if Revit 2022 or newer and Pipe Accessory schedule
    if is_revit_2022_or_newer and not __shiftclick__ and family_field is not None:
        # Create a schedule filter using EndsWith on the Family field
        schedule_filter = ScheduleFilter(
            family_field.FieldId,
            ScheduleFilterType.EndsWith,
            "WS"
        )
        definition.AddFilter(schedule_filter)
        # print("Added filter for family names ending with 'WS'")

    t.Commit()
else:
    print("'{}' already exists.".format(schedule_name))