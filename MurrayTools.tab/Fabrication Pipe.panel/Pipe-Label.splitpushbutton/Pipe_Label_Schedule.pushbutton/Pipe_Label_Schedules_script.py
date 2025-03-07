import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')

from System.Collections.Generic import List
from Autodesk.Revit.DB import BuiltInCategory, Transaction, ElementId, ViewSchedule, FilteredElementCollector, ParameterElement, ScheduleFieldType, BuiltInParameter, ParameterFilterElement, ParameterFilterRuleFactory

# Define the active Revit application and document
doc = __revit__.ActiveUIDocument.Document

# Function to check if a schedule with a specific name and category exists
def schedule_exists(schedule_name, category_id):
    schedules_collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    for schedule in schedules_collector:
        if schedule.Name == schedule_name and schedule.Definition.CategoryId == category_id:
            return True
    return False

# Define the fields for the schedule and their desired order
fieldNames = [
    ("Family", "Description"), #This column will be hidden
    ("FP_Product Entry", "LABEL SIZE"),
    ("FP_Service Abbreviation", "SYSTEM ABBR."),
    ("FP_Service Name", "SERVICE NAME")
]

# Get the category id for the elements you want in the schedule
categoryId = ElementId(BuiltInCategory.OST_PipeAccessory)

# Check if the schedule already exists
schedule_name = "PIPE LABEL SCHEDULE"
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

    # Insert the "Count" field
    countField = definition.AddField(ScheduleFieldType.Count)
    countField.ColumnHeading = "Qty"

    # Function to add field by name
    def add_field_by_name(definition, paramName, userColumnName, parameters):
        # print paramName
        # print paramName == "Family"
        if paramName == "Family":
            # Handle the Family parameter separately using the built-in parameter
            paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
            field = definition.AddField(ScheduleFieldType.Instance, paramId)
            field.ColumnHeading = userColumnName
            field.IsHidden = True
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

    # Add fields to the schedule in the specified order
    for paramName, userColumnName in fieldNames:
        add_field_by_name(definition, paramName, userColumnName, parameters)

    t.Commit()
else:
    print("'{}' already exists.".format(schedule_name))

