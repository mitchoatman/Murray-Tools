# -*- coding: UTF-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FabricationPart, ElementId
from pyrevit import script
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString
import sys
doc = __revit__.ActiveUIDocument.Document
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
# Creating a collector instance and collecting all the Fabrication Pipework
collector = FilteredElementCollector(doc) \
                    .OfCategory(BuiltInCategory.OST_FabricationPipework) \
                    .WhereElementIsNotElementType()
# Filter for FabricationPart with CID 2041 and straight pipes
Pipe_collector = [
    elem for elem in collector
    if isinstance(elem, FabricationPart) and elem.ItemCustomId == 2041 and getattr(elem, "IsAStraight", False)
]
if not Pipe_collector:
    print("No straight fabrication pipes with CID 2041 found.")
    sys.exit()
def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()
# Dictionary to store total length for each level, material, and system
level_material_total_lengths = {}
level_total_lengths = {}
material_total_lengths = {}
system_level_total_lengths = {}
level_system_material_total_lengths = {}
total_length = 0.0
# Iterate over pipes and collect Length data
for pipe in Pipe_collector:
    len_param = pipe.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH)
    if len_param and len_param.HasValue:
        length = len_param.AsDouble()
        total_length += length
        level_id = pipe.LevelId
        level_name = doc.GetElement(level_id).Name
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
# Table 1: Total Lengths by Level
level_table_data = [[level, "{:.2f}".format(length)] for level, length in level_total_lengths.items()]
level_table_data.append(["Total", "{:.2f}".format(total_length)])
# Table 2: Total Lengths by Level and Material
level_material_table_data = []
for level, materials in level_material_total_lengths.items():
    for material, length in materials.items():
        level_material_table_data.append([level, material, "{:.2f}".format(length)])
# Table 3: Total Lengths by Material
material_table_data = [[material, "{:.2f}".format(length)] for material, length in material_total_lengths.items()]
# Table 4: Total Lengths by Level, System, and Material
level_system_material_table_data = []
for level, systems in level_system_material_total_lengths.items():
    for system, materials in systems.items():
        for material, length in materials.items():
            level_system_material_table_data.append([level, system, material, "{:.2f}".format(length)])
# Sort the table data
material_table_data = sorted(material_table_data, key=lambda x: x[0])
level_material_table_data = sorted(level_material_table_data, key=lambda x: (x[0], x[1]))
level_system_material_table_data = sorted(level_system_material_table_data, key=lambda x: (x[0], x[1], x[2]))
# Print tables with left-aligned columns
output.print_table(table_data=material_table_data,
                   title="Total Lengths of Fabrication Pipes by Material",
                   columns=["Material", "Length (Linear Feet)"],
                   formats=['', '{}'])
# output.print_table(table_data=level_table_data,
                   # title="Total Lengths of Fabrication Pipes by Level",
                   # columns=["Level", "Length (Linear Feet)"],
                   # formats=['', '{}'])
output.print_table(table_data=level_material_table_data,
                   title="Total Lengths of Fabrication Pipes by Level and Material",
                   columns=["Level", "Material", "Length (Linear Feet)"],
                   formats=['', '', '{}'])
output.print_table(table_data=level_system_material_table_data,
                   title="Total Lengths of Fabrication Pipes by Level, System, and Material",
                   columns=["Level", "System", "Material", "Length (Linear Feet)"],
                   formats=['', '', '', '{}'])
# Collect fabrication hangers
hanger_collector = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_FabricationHangers) \
    .WhereElementIsNotElementType()
hanger_parts = [
    elem for elem in hanger_collector
    if isinstance(elem, FabricationPart) and elem.ItemCustomId == 838
]
# Dictionary to store counts for each level, system, and hanger name
level_system_hanger_counts = {}
if hanger_parts:
    for hanger in hanger_parts:
        level_id = hanger.LevelId
        if level_id != ElementId.InvalidElementId:
            level_name = doc.GetElement(level_id).Name
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
# Print hanger table
if hanger_table_data:
    output.print_table(table_data=hanger_table_data,
                       title="Total Counts of Fabrication Hangers by Level, System, and Hanger Type",
                       columns=["Level", "System", "Hanger Type", "Count"],
                       formats=['', '', '', '{}'])
else:
    output.print_md("**No fabrication hangers with CID 838 found.**")