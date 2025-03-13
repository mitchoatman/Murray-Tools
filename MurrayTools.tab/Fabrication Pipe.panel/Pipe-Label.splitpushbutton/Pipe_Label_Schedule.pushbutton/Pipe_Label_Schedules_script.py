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
    ("Family", "Description"), #This column will be hidden
    ("FP_Product Entry", "LABEL SIZE"),
    ("FP_Service Abbreviation", "SYSTEM ABBR."),
    ("FP_Service Name", "SERVICE NAME")
]

# Get the category id for the elements you want in the schedule
categoryId = ElementId(BuiltInCategory.OST_PipeAccessory)

# Check Revit version
revit_version = int(app.VersionNumber)
is_revit_2022_or_newer = revit_version >= 2022

# Find the ElementId of the "Pipe Label" family type
pipe_label_type_id = None
# print("DEBUG: Finding 'Pipe Label' family type ID:")
pipe_accessories = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
for elem in pipe_accessories:
    fam_param = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
    if fam_param and fam_param.AsValueString() == "Pipe Label":
        pipe_label_type_id = elem.GetTypeId()
        # print("Found 'Pipe Label' with TypeId: {}".format(pipe_label_type_id.IntegerValue))
        break  # We only need one instance to get the type ID

if pipe_label_type_id is None:
    print("WARNING: No 'Pipe Label' family found in the model!")

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
        return field  # Return the field for potential use in filtering

    # Add fields to the schedule and store the family field
    family_field = None
    for paramName, userColumnName in fieldNames:
        field = add_field_by_name(definition, paramName, userColumnName, parameters)
        if paramName == "Family":
            family_field = field

    # Add filter using Family and Type ID if Revit 2022 or newer
    if is_revit_2022_or_newer and pipe_label_type_id is not None:
        try:
            # Add a field for Family and Type (not just Type)
            type_field = definition.AddField(ScheduleFieldType.Instance, ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM))
            type_field.ColumnHeading = "Family and Type"
            type_field.IsHidden = True  # Hide this field since we only need it for filtering
            
            # Create a schedule filter using Equals on the Family and Type ID
            schedule_filter = ScheduleFilter(
                type_field.FieldId,
                ScheduleFilterType.Equal,
                pipe_label_type_id
            )
            
            # Add the filter to the schedule definition
            definition.AddFilter(schedule_filter)
            # print("Filter added successfully for 'Pipe Label' using TypeId: {}".format(pipe_label_type_id.IntegerValue))
        except Exception as e:
            print("Failed to add filter: {}".format(str(e)))
            print("Note: Could not filter by Family and Type ID")

    t.Commit()
else:
    print("'{}' already exists.".format(schedule_name))