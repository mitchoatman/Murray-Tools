import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')

from System.Collections.Generic import List
from SharedParam.Add_Parameters import Shared_Params
from Autodesk.Revit.DB import (
    BuiltInCategory, Transaction, ElementId, ViewSchedule, 
    FilteredElementCollector, ParameterElement, ScheduleFieldType, 
    BuiltInParameter, ScheduleSortGroupField, ScheduleSortOrder,
    FormatOptions
)

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
    ("Count", "Qty"),
    ("FP_Centerline Length", "Length"),
    ("FP_Product Entry", "Size"),
    ("Family", "Description"),
    ("Pipe Material Name", "Material"),
    ("Description BOM Suffix", "Total")
]

# Get the category id for MEP Fabrication Pipework
categoryId = ElementId(BuiltInCategory.OST_FabricationPipework)

# Check if the schedule already exists
schedule_name = "MEP PIPEWORK SCHEDULE"
if not schedule_exists(schedule_name, categoryId):
    # Start a new transaction
    t = Transaction(doc, "Create MEP Pipework Schedule")
    t.Start()

    # Create the schedule
    schedule = ViewSchedule.CreateSchedule(doc, categoryId)

    # Set the schedule name
    schedule.Name = schedule_name

    # Get the ScheduleDefinition from the ViewSchedule
    definition = schedule.Definition

    # Turn off "Itemized every instance"
    definition.IsItemized = False

    # Get all parameters in the document
    parameters = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

    # Dictionary to store ScheduleFieldIds for sorting/grouping
    field_ids = {}

    def add_field_by_name(definition, paramName, userColumnName, parameters):
        try:
            if paramName == "Count":
                field = definition.AddField(ScheduleFieldType.Count)
                field.ColumnHeading = userColumnName
                field_ids[paramName] = field.FieldId
                # print("Added field '{}' ({})".format(paramName, userColumnName))
            elif paramName == "Family":
                paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
                field = definition.AddField(ScheduleFieldType.Instance, paramId)
                field.ColumnHeading = userColumnName
                field_ids[paramName] = field.FieldId
                # print("Added field '{}' ({})".format(paramName, userColumnName))
            elif paramName == "Pipe Material Name":
                paramId = ElementId(BuiltInParameter.FABRICATION_PART_MATERIAL)
                field = definition.AddField(ScheduleFieldType.Instance, paramId)
                field.ColumnHeading = userColumnName
                field_ids[paramName] = field.FieldId
                # print("Added field '{}' ({})".format(paramName, userColumnName))
            else:
                # Attempt to find a parameter with matching name
                parameter = next((p for p in parameters if p.Name == paramName), None)
                if parameter is not None:
                    paramId = parameter.Id
                    field = definition.AddField(ScheduleFieldType.Instance, paramId)
                    field.ColumnHeading = userColumnName
                    field_ids[paramName] = field.FieldId
                    # Configure FP_Centerline Length: Suppress 0 feet
                    if paramName == "FP_Centerline Length":
                        try:
                            format_options = field.GetFormatOptions()
                            format_options.UseDefaultFormatting = False
                            format_options.SuppressZeroFeet = True
                            # Example additional formatting options
                            format_options.Accuracy = 0.01  # Set precision to 1/100 inch
                            format_options.SuppressTrailingZeros = True  # Remove trailing zeros (e.g., 6.00" to 6")
                            format_options.SuppressSpaces = True  # Remove spaces (e.g., 1' 6" to 1'6")
                            field.SetFormatOptions(format_options)
                        except Exception as e:
                            print("Failed to apply formatting for field '{}' ({}): {}".format(paramName, userColumnName, str(e)))
                else:
                    print("Parameter '{}' not found in document.".format(paramName))
        except Exception as e:
            print("Failed to add field '{}' ({}): {}".format(paramName, userColumnName, str(e)))

    # Add fields to the schedule in the specified order
    for paramName, userColumnName in fieldNames:
        add_field_by_name(definition, paramName, userColumnName, parameters)

    # Add sorting and grouping by Family, Size, and Material
    sort_group_fields = [
        ("Family", "Description"),
        ("FP_Product Entry", "Size"),
        ("Pipe Material Name", "Material")
    ]

    for paramName, userColumnName in sort_group_fields:
        if paramName in field_ids:
            sort_field = ScheduleSortGroupField(field_ids[paramName])
            sort_field.SortOrder = ScheduleSortOrder.Ascending
            sort_field.ShowHeader = False  # Disable grouping header
            definition.AddSortGroupField(sort_field)
        else:
            print("Cannot sort/group by '{}' ({}) - field not found.".format(paramName, userColumnName))

    t.Commit()
else:
    print("Schedule '{}' already exists.".format(schedule_name))