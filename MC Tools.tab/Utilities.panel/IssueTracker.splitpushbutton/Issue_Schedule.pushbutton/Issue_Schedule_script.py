import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')

from System.Collections.Generic import List
from Autodesk.Revit.DB import BuiltInCategory, Transaction, ElementId, ViewSchedule, FilteredElementCollector, ParameterElement, ScheduleFieldType, BuiltInParameter, ScheduleFilter, ScheduleFilterType

# Define the active Revit application and document
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

# Function to check if a schedule with a specific name and category exists
def schedule_exists(schedule_name, category_id):
    schedules_collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    for schedule in schedules_collector:
        if schedule.Name == schedule_name and schedule.Definition.CategoryId == category_id:
            return True
    return False

# Define the fields for the schedule and their desired order
fieldNames = [
    ("Level", "LEVEL"),
    ("Family", "FAMILY"),
    ("Comments", "COMMENTS")
]

# Get the category id for the elements you want in the schedule
categoryId = ElementId(BuiltInCategory.OST_GenericModel)

# Check if the schedule already exists
schedule_name = "MC-ISSUE SCHEDULE"
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

    # Function to add field by name
    def add_field_by_name(definition, paramName, userColumnName):
        if paramName == "Family":
            paramId = ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
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
        return field

    # Add fields to the schedule and store the comments field
    comments_field = None
    for paramName, userColumnName in fieldNames:
        field = add_field_by_name(definition, paramName, userColumnName)
        if paramName == "Comments":
            comments_field = field

    # Add filter for Comments not empty
    if comments_field:
        comments_filter = ScheduleFilter(comments_field.FieldId, ScheduleFilterType.GreaterThan, "")
        definition.AddFilter(comments_filter)

    t.Commit()
    # print("'{}' created successfully.".format(schedule_name))
else:
    print("'{}' already exists.".format(schedule_name))