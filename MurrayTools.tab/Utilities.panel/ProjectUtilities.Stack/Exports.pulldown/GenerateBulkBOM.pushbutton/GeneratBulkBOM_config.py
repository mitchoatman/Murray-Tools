import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')

from Autodesk.Revit.DB import (
    BuiltInCategory, Transaction, ElementId, ViewSchedule, 
    FilteredElementCollector, ParameterElement, ScheduleFieldType, 
    BuiltInParameter, ScheduleSortGroupField, ScheduleSortOrder,
    ScheduleFilter, ScheduleFilterType, ScheduleFieldDisplayType
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

# Define the fields for non-hanger schedules
fieldNames_default = [
    ("Count", "Qty"),
    ("FP_Centerline Length", "Length"),
    ("FP_Product Entry", "Size"),
    ("Family", "Description"),
    ("FP_Part Material", "Material"),
    ("FP_Unit Cost", "Unit Cost"),
    ("FP_Total", "Total"),
    ("FP_Service Type", "FP_Service Type")
]

# Define the fields for hanger schedule
fieldNames_hangers = [
    ("Count", "Qty"),
    ("FP_Bearer Length", "Length"),
    ("FP_Product Entry", "Size"),
    ("Family", "Description"),
    ("FP_Part Material", "Material"),
    ("FP_Unit Cost", "Unit Cost"),
    ("FP_Total", "Total"),
    ("FP_Service Type", "FP_Service Type")
]

# Define the fields for hanger schedule
fieldNames_hanger_rod = [
    ("Count", "Qty"),
    ("FP_Rod Length A", "Length"),
    ("FP_Rod Length B", "Length"),
    #("Total Rod Length", "Total Rod Length"),
    ("FP_Rod Size", "Size"),
    ("Family", "Description"),
    ("FP_Part Material", "Material"),
    ("FP_Unit Cost", "Unit Cost"),
    ("FP_Total", "Total"),
    ("FP_Service Type", "FP_Service Type")
]

# Get the category ids
pipework_category_id = ElementId(BuiltInCategory.OST_FabricationPipework)
hanger_category_id = ElementId(BuiltInCategory.OST_FabricationHangers)

# Function to create a schedule with specified name, category, fields, and filters
def create_schedule(schedule_name, category_id, filters, fieldNames):
    if not schedule_exists(schedule_name, category_id):
        # Start a new transaction
        t = Transaction(doc, "Create {}".format(schedule_name))
        t.Start()

        # Create the schedule
        schedule = ViewSchedule.CreateSchedule(doc, category_id)

        # Set the schedule name
        schedule.Name = schedule_name

        # Get the ScheduleDefinition from the ViewSchedule
        definition = schedule.Definition

        # Turn off "Itemized every instance"
        definition.IsItemized = False

        # Get all parameters in the document
        parameters = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

        # Dictionary to store ScheduleFieldIds for sorting/grouping and filtering
        field_ids = {}

        def add_field_by_name(definition, paramName, userColumnName, parameters):
            try:
                if paramName == "Count":
                    field = definition.AddField(ScheduleFieldType.Count)
                    field.ColumnHeading = userColumnName
                    field_ids[paramName] = field.FieldId

                elif paramName == "Family":
                    paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
                    field = definition.AddField(ScheduleFieldType.Instance, paramId)
                    field.ColumnHeading = userColumnName
                    field_ids[paramName] = field.FieldId

                # elif paramName == "Total Rod Length":
                    # try:
                        # # Correct overload: AddField(ScheduleFieldType) with no second argument for formula fields
                        # field = definition.AddField(ScheduleFieldType.Formula)
                        # field.ColumnHeading = userColumnName
                        
                        # # Set the formula use double quotes around parameter names (most reliable via API)
                        # # Ensure exact spelling/case matches your shared parameters
                        # field.SetFormula('"FP_Rod Length A" + "FP_Rod Length B"')
                        
                        # # Enable totals display (grand total at the bottom of the schedule)
                        # field.DisplayType = ScheduleFieldDisplayType.Totals
                        
                        # field_ids[paramName] = field.FieldId
                        # print("Successfully added calculated field: Total Rod Length with formula '{}'".format(field.GetFormula()))
                    # except Exception as ex:
                        # print("Failed to add calculated field 'Total Rod Length': {}".format(str(ex)))

                else:
                    # Shared/project parameters
                    parameter = next((p for p in parameters if p.Name == paramName), None)
                    if parameter is not None:
                        paramId = parameter.Id
                        field = definition.AddField(ScheduleFieldType.Instance, paramId)
                        field.ColumnHeading = userColumnName
                        field_ids[paramName] = field.FieldId

                        # Enable totals for length fields
                        if paramName in ("FP_Centerline Length", "FP_Bearer Length", "FP_Rod Length A", "FP_Rod Length B", "Total Rod Length"):
                            field.DisplayType = ScheduleFieldDisplayType.Totals

                        # Hide internal/helper fields
                        if paramName in ("FP_Service Type", "FP_Rod Length A", "FP_Rod Length B"):
                            try:
                                field.IsHidden = True
                            except:
                                pass  # Some versions restrict IsHidden timing
                    else:
                        print("Parameter not found: '{}' in schedule '{}'".format(paramName, schedule_name))

            except Exception as ex:
                print("Failed to add field '{}': {}".format(paramName, str(ex)))

        # Add fields to the schedule in the specified order
        for paramName, userColumnName in fieldNames:
            add_field_by_name(definition, paramName, userColumnName, parameters)

        # Sorting and grouping apply different rules for hanger rod schedule
        if schedule_name == "MEP FAB HANGER ROD":
            # Hanger rod: Group by FP_Rod Size first, then Family; footer with totals only on rod size
            sort_group_fields = [
                ("FP_Rod Size", "Size"),
                ("Family", "Description")
            ]

            for index, (paramName, userColumnName) in enumerate(sort_group_fields):
                if paramName in field_ids:
                    sort_field = ScheduleSortGroupField(field_ids[paramName])
                    sort_field.SortOrder = ScheduleSortOrder.Ascending
                    sort_field.ShowHeader = False

                    if index == 0:  # Primary group: FP_Rod Size
                        sort_field.ShowFooter = True
                        sort_field.ShowFooterTitle = True
                        sort_field.ShowFooterCount = False
                    else:
                        sort_field.ShowFooter = False
                    definition.AddSortGroupField(sort_field)

        else:
            # All other schedules: original grouping (Family Size Material)
            sort_group_fields = [
                ("Family", "Description"),
                ("FP_Product Entry", "Size"),
                ("FP_Part Material", "Material")
            ]

            for paramName, userColumnName in sort_group_fields:
                if paramName in field_ids:
                    sort_field = ScheduleSortGroupField(field_ids[paramName])
                    sort_field.SortOrder = ScheduleSortOrder.Ascending
                    sort_field.ShowHeader = False
                    definition.AddSortGroupField(sort_field)

        # Apply filters
        try:
            for filter_field, filter_type, filter_value in filters:
                if filter_field in field_ids:
                    filter_obj = ScheduleFilter(
                        field_ids[filter_field],
                        filter_type,
                        filter_value
                    )
                    definition.AddFilter(filter_obj)
        except Exception:
            pass

        t.Commit()
    else:
        print("Schedule '{}' already exists.".format(schedule_name))

# Define filters for each schedule
pipework_filters = [
    ("FP_Part Material", ScheduleFilterType.NotContains, "Brass"),
    ("FP_Part Material", ScheduleFilterType.NotContains, "Copper"),
    ("FP_Part Material", ScheduleFilterType.NotContains, "Bronze"),
    ("FP_Service Type", ScheduleFilterType.Equal, "Pipework")
]

valve_filters = [
    ("FP_Service Type", ScheduleFilterType.Equal, "Valve")
]

hanger_filters = [
    ("FP_Service Type", ScheduleFilterType.Equal, "Hanger")
]

copper_filters = [
    ("FP_Part Material", ScheduleFilterType.NotContains, "Iron"),
    ("FP_Part Material", ScheduleFilterType.NotContains, "Steel"),
    ("FP_Part Material", ScheduleFilterType.NotContains, "PVC"),
    ("FP_Part Material", ScheduleFilterType.NotContains, "Poly"),
    ("FP_Service Type", ScheduleFilterType.Equal, "Pipework")
]

# Create all schedules with appropriate field lists
create_schedule("MEP FAB PIPE", pipework_category_id, pipework_filters, fieldNames_default)
create_schedule("MEP FAB VALVES", pipework_category_id, valve_filters, fieldNames_default)
create_schedule("MEP FAB HANGERS", hanger_category_id, hanger_filters, fieldNames_hangers)
create_schedule("MEP FAB COPPER PIPE", pipework_category_id, copper_filters, fieldNames_default)
create_schedule("MEP FAB HANGER ROD", hanger_category_id, hanger_filters, fieldNames_hanger_rod)