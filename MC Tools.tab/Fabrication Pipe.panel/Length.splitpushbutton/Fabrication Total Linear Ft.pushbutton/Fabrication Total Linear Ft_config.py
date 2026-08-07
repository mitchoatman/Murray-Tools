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

# Get currently selected element IDs from the active UI document
selected_ids = uidoc.Selection.GetElementIds()

if not selected_ids:
    output.print_md("**No elements selected. Please select fabrication parts in the active view and re-run the script.**")
    sys.exit()

# Fetch selected elements into objects
selected_elements = [doc.GetElement(elem_id) for elem_id in selected_ids]

# Filter selected elements for FabricationPart with CID 2041 and straight pipes
Pipe_collector = [
    elem for elem in selected_elements
    if isinstance(elem, FabricationPart) and elem.ItemCustomId == 2041 and getattr(elem, "IsAStraight", False)
]

def get_parameter_value_by_name(element, parameterName):
    param = element.LookupParameter(parameterName)
    return param.AsValueString() if param else ""

# Dictionaries to store total length for each level, material, and system
level_material_total_lengths = {}
level_total_lengths = {}
material_total_lengths = {}
system_level_total_lengths = {}
level_system_material_total_lengths = {}
total_length = 0.0

# Iterate over selected pipes and collect Length data
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

# Prepare data for tables
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

# Sort the table data
material_table_data = sorted(material_table_data, key=lambda x: x[0])
level_material_table_data = sorted(level_material_table_data, key=lambda x: (x[0], x[1]))
level_system_material_table_data = sorted(level_system_material_table_data, key=lambda x: (x[0], x[1], x[2]))

# Print pipe tables if data exists
if material_table_data:
    output.print_table(table_data=material_table_data,
                        title="Total Lengths of Selected Fabrication Pipes by Material",
                        columns=["Material", "Length (Linear Feet)"],
                        formats=['', '{}'])

    output.print_table(table_data=level_material_table_data,
                        title="Total Lengths of Selected Fabrication Pipes by Level and Material",
                        columns=["Level", "Material", "Length (Linear Feet)"],
                        formats=['', '', '{}'])

    output.print_table(table_data=level_system_material_table_data,
                        title="Total Lengths of Selected Fabrication Pipes by Level, System, and Material",
                        columns=["Level", "System", "Material", "Length (Linear Feet)"],
                        formats=['', '', '', '{}'])
else:
    output.print_md("**No straight fabrication pipes with CID 2041 found in the current selection.**")

# Filter selected elements for Fabrication Hangers with CID 838
hanger_parts = [
    elem for elem in selected_elements
    if isinstance(elem, FabricationPart) and elem.ItemCustomId == 838
]

# Dictionary to store counts for each level, system, and hanger name
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
                       title="Total Counts of Selected Fabrication Hangers by Level, System, and Hanger Type",
                       columns=["Level", "System", "Hanger Type", "Count"],
                       formats=['', '', '', '{}'])