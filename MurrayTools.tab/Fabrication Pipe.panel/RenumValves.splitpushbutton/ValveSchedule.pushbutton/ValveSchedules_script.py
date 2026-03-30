import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    BuiltInCategory, Transaction, ElementId, ViewSchedule, 
    FilteredElementCollector, ParameterElement, ScheduleFieldType, 
    BuiltInParameter, ScheduleFilter, ScheduleFilterType
)

from Parameters.Add_SharedParameters import Shared_Params
Shared_Params()

from Parameters.FabPart_Params import Sync_FP_Params_Entire_Model
Sync_FP_Params_Entire_Model()

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
    ("FP_Valve Number", "Valve Number"),
    ("Family", "Description"),
    ("Reference Level", "Level"),
    ("FP_Service Type", "FP_Service Type")
]

# Get the category id for the elements you want in the schedule
categoryId = ElementId(BuiltInCategory.OST_FabricationPipework)

# Check if the schedule already exists
schedule_name = "VALVE SCHEDULE"
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

    # Dictionary to store ScheduleFieldIds for filtering
    field_ids = {}

    # Function to add field by name
    def add_field_by_name(definition, paramName, userColumnName, parameters):
        if paramName == "Family":
            # Handle the Family parameter separately using the built-in parameter
            paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
            field = definition.AddField(ScheduleFieldType.Instance, paramId)
            field.ColumnHeading = userColumnName
            field_ids[paramName] = field.FieldId
        elif paramName == "Reference Level":
            # Handle the Reference Level parameter using the built-in parameter
            paramId = ElementId(BuiltInParameter.FABRICATION_LEVEL_PARAM)
            field = definition.AddField(ScheduleFieldType.Instance, paramId)
            field.ColumnHeading = userColumnName
            field_ids[paramName] = field.FieldId
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
                field_ids[paramName] = field.FieldId
                # Set FP_Service Type as hidden
                if paramName == "FP_Service Type":
                    try:
                        field.IsHidden = True
                    except Exception:
                        pass

    # Add fields to the schedule in the specified order
    for paramName, userColumnName in fieldNames:
        add_field_by_name(definition, paramName, userColumnName, parameters)

    # Add filter for FP_Service Type = Valve
    try:
        if "FP_Service Type" in field_ids:
            filter_obj = ScheduleFilter(
                field_ids["FP_Service Type"],
                ScheduleFilterType.Equal,
                "Valve"
            )
            definition.AddFilter(filter_obj)
    except Exception:
        pass

    t.Commit()
else:
    print("Schedule '{}' already exists.".format(schedule_name))