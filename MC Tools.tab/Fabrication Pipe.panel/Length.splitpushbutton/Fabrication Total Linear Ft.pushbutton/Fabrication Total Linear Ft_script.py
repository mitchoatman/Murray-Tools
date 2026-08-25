# -*- coding: UTF-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FabricationPart, ElementId
from pyrevit import script
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString
import sys

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# Get project name from Project Information and file name from doc.Title
project_info_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectInformation).ToElements()
project_name = None
if project_info_collector:
    project_info = project_info_collector[0]
    project_name_param = project_info.LookupParameter("Project Name")
    if project_name_param and project_name_param.HasValue:
        project_name = project_name_param.AsString()
project_name = project_name or "Untitled Project"
file_name = doc.Title or "Untitled File"

# Print project name and file name as a title
output.print_md("# {} - {}".format(project_name, file_name))

# Check for user selection
selected_ids = uidoc.Selection.GetElementIds()

if selected_ids:
    # Selection Mode: Get elements from current selection
    elements_to_process = [doc.GetElement(elem_id) for elem_id in selected_ids if doc.GetElement(elem_id) is not None]
    mode_label = "Selected "
else:
    # Global Mode: Collect all fabrication pipework and hangers from the entire document
    pipe_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationPipework).WhereElementIsNotElementType()
    hanger_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangers).WhereElementIsNotElementType()
    elements_to_process = list(pipe_collector) + list(hanger_collector)
    mode_label = ""

# Helper function for parameter retrieval
def get_parameter_value_by_name(element, parameterName):
    param = element.LookupParameter(parameterName)
    return param.AsValueString() if param and param.HasValue else ""

# -------------------------------------------------------------
# 1. PROCESS PIPES (CID 2041, Straight Pipes)
# -------------------------------------------------------------
Pipe_collector = [
    elem for elem in elements_to_process
    if isinstance(elem, FabricationPart) and elem.ItemCustomId == 2041 and getattr(elem, "IsAStraight", False)
]

level_material_total_lengths = {}
level_total_lengths = {}
material_total_lengths = {}
system_level_total_lengths = {}
level_system_material_total_lengths = {}
total_length = 0.0

for pipe in Pipe_collector:
    len_param = pipe.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH)
    if len_param and len_param.HasValue:
        length = len_param.AsDouble()
        total_length += length
        level_id = pipe.LevelId
        
        level_elem = doc.GetElement(level_id)
        level_name = level_elem.Name if level_elem else "Unassigned Level"
        
        material_name = get_parameter_value_by_name(pipe, 'Part Material')
        system_name = get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Name')
        
        # Update level totals
        if level_name not in level_total_lengths:
            level_total_lengths[level_name] = 0.0
        level_total_lengths[level_name] += length
        
        # Update level and material totals
        if level_name not in level_material_total_lengths:
            level_material_total_lengths[level_name] = {}
        if material_name not in level_material_total_lengths[level_name]:
            level_material_total_lengths[level_name][material_name] = 0.0
        level_material_total_lengths[level_name][material_name] += length
        
        # Update material totals
        if material_name not in material_total_lengths:
            material_total_lengths[material_name] = 0.0
        material_total_lengths[material_name] += length
        
        # Update level and system totals
        if level_name not in system_level_total_lengths:
            system_level_total_lengths[level_name] = {}
        if system_name not in system_level_total_lengths[level_name]:
            system_level_total_lengths[level_name][system_name] = 0.0
        system_level_total_lengths[level_name][system_name] += length
        
        # Update level, system, and material totals
        if level_name not in level_system_material_total_lengths:
            level_system_material_total_lengths[level_name] = {}
        if system_name not in level_system_material_total_lengths[level_name]:
            level_system_material_total_lengths[level_name][system_name] = {}
        if material_name not in level_system_material_total_lengths[level_name][system_name]:
            level_system_material_total_lengths[level_name][system_name][material_name] = 0.0
        level_system_material_total_lengths[level_name][system_name][material_name] += length

# Prepare pipe table data
material_table_data = [[material, "{:.2f}".format(length)] for material, length in material_total_lengths.items()]

level_material_table_data = []
for level, materials in level_material_total_lengths.items():
    for material, length in materials.items():
        level_material_table_data.append([level, material, "{:.2f}".format(length)])

level_system_material_table_data = []
for level, systems in level_system_material_total_lengths.items():
    for system, materials in systems.items():
        for material, length in materials.items():
            level_system_material_table_data.append([level, system, material, "{:.2f}".format(length)])

# Sort pipe tables
material_table_data = sorted(material_table_data, key=lambda x: x[0])
level_material_table_data = sorted(level_material_table_data, key=lambda x: (x[0], x[1]))
level_system_material_table_data = sorted(level_system_material_table_data, key=lambda x: (x[0], x[1], x[2]))

# Print pipe tables if data exists
if material_table_data:
    output.print_table(table_data=material_table_data,
                       title="Total Lengths of {}Fabrication Pipes by Material".format(mode_label),
                       columns=["Material", "Length (Linear Feet)"],
                       formats=['', '{}'])

    output.print_table(table_data=level_material_table_data,
                       title="Total Lengths of {}Fabrication Pipes by Level and Material".format(mode_label),
                       columns=["Level", "Material", "Length (Linear Feet)"],
                       formats=['', '', '{}'])

    output.print_table(table_data=level_system_material_table_data,
                       title="Total Lengths of {}Fabrication Pipes by Level, System, and Material".format(mode_label),
                       columns=["Level", "System", "Material", "Length (Linear Feet)"],
                       formats=['', '', '', '{}'])
else:
    output.print_md("**No straight fabrication pipes with CID 2041 found in the {}scope.**".format("current selection" if selected_ids else "entire project"))


# -------------------------------------------------------------
# 2. PROCESS HANGERS (CID 838)
# -------------------------------------------------------------
hanger_parts = [
    elem for elem in elements_to_process
    if isinstance(elem, FabricationPart) and elem.ItemCustomId == 838
]

level_system_hanger_counts = {}
if hanger_parts:
    for hanger in hanger_parts:
        level_id = hanger.LevelId
        if level_id != ElementId.InvalidElementId:
            level_elem = doc.GetElement(level_id)
            level_name = level_elem.Name if level_elem else "Unassigned Level"
            system_name = get_parameter_value_by_name_AsString(hanger, 'Fabrication Service Name')
            hanger_name = get_parameter_value_by_name(hanger, 'Family')
            
            if level_name not in level_system_hanger_counts:
                level_system_hanger_counts[level_name] = {}
            if system_name not in level_system_hanger_counts[level_name]:
                level_system_hanger_counts[level_name][system_name] = {}
            if hanger_name not in level_system_hanger_counts[level_name][system_name]:
                level_system_hanger_counts[level_name][system_name][hanger_name] = 0
            level_system_hanger_counts[level_name][system_name][hanger_name] += 1

# Prepare data for hanger table
hanger_table_data = []
for level, systems in level_system_hanger_counts.items():
    for system, hangers in systems.items():
        for hanger_name, count in hangers.items():
            hanger_table_data.append([level, system, hanger_name, str(count)])

# Sort the hanger table data
hanger_table_data = sorted(hanger_table_data, key=lambda x: (x[0], x[1], x[2]))

# Print hanger table if data exists
if hanger_table_data:
    output.print_table(table_data=hanger_table_data,
                       title="Total Counts of {}Fabrication Hangers by Level, System, and Hanger Type".format(mode_label),
                       columns=["Level", "System", "Hanger Type", "Count"],
                       formats=['', '', '', '{}'])
else:
    output.print_md("**No fabrication hangers with CID 838 found in the {}scope.**".format("current selection" if selected_ids else "entire project"))