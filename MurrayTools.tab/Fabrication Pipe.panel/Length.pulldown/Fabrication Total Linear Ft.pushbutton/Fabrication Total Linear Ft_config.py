
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter

doc = __revit__.ActiveUIDocument.Document

# Creating a collector instance and collecting all the Fabrication Pipe and Fittings
Pipe_collector = FilteredElementCollector(doc) \
                    .OfCategory(BuiltInCategory.OST_FabricationPipework) \
                    .WhereElementIsNotElementType()

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()

# Dictionary to store total length for each level and material
level_material_total_lengths = {}

level_total_lengths = {}

# Iterate over pipes and collect Length data
for pipe in Pipe_collector:
    if pipe.IsAStraight:
        len_param = pipe.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH)
        if len_param:
            length = len_param.AsDouble()
            level_id = pipe.LevelId
            level_name = doc.GetElement(level_id).Name
            material_name = get_parameter_value_by_name(pipe, 'Part Material')
            
            if level_name not in level_total_lengths:
                level_total_lengths[level_name] = 0.0
            level_total_lengths[level_name] += length
            
            if level_name not in level_material_total_lengths:
                level_material_total_lengths[level_name] = {}
            
            if material_name not in level_material_total_lengths[level_name]:
                level_material_total_lengths[level_name][material_name] = 0.0
            
            level_material_total_lengths[level_name][material_name] += length

# Now that results are collected, print the total for each level and material
print("Total Lengths of Fabrication Pipes by Level:")
for level, length in level_total_lengths.items():
    print("  {}: {} Linear feet".format(level, length))
print("----------\nTotal Lengths of Fabrication Pipes by Level and Material:")
for level, materials in level_material_total_lengths.items():
    print("\nLevel: {}".format(level))
    for material, length in materials.items():
        print("  Material: [{}:] {} Linear feet".format(material, length))





# from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter

# doc = __revit__.ActiveUIDocument.Document

# # Creating a collector instance and collecting all the Fabrication Pipe and Fittings
# Pipe_collector = FilteredElementCollector(doc) \
                    # .OfCategory(BuiltInCategory.OST_FabricationPipework) \
                    # .WhereElementIsNotElementType()

# def get_parameter_value_by_name(element, parameterName):
    # return element.LookupParameter(parameterName).AsValueString()

# # Dictionary to store total length for each level
# level_total_lengths = {}

# material_total_lengths = {}

# total_length = 0.0

# # Iterate over pipes and collect Length data
# for pipe in Pipe_collector:
    # if pipe.IsAStraight:
        # len_param = pipe.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH)
        # if len_param:
            # total_length = total_length + len_param.AsDouble()
            # level_id = pipe.LevelId
            # level_name = doc.GetElement(level_id).Name
            # material_name = get_parameter_value_by_name(pipe, 'Part Material')
            # if level_name not in level_total_lengths:
                # level_total_lengths[level_name] = 0.0
            # level_total_lengths[level_name] += len_param.AsDouble()
            # if material_name not in material_total_lengths or level_name not in level_total_lengths:
                # material_total_lengths[material_name] = 0.0
                # level_total_lengths[level_name] = 0.0
            # level_total_lengths[level_name] += len_param.AsDouble()
            # material_total_lengths[material_name] += len_param.AsDouble()

# # Now that results are collected, print the total for each level
# print("Total Lengths of Fabrication Pipes by Level and Material:")
# for level, length in level_total_lengths.items():
    # print("  {}: {} Linear feet \n----------".format(level, length))
# for material, length in material_total_lengths.items():
    # print("  {}:\n {}:\n {} Linear feet \n----------".format(level, material, length))
# # now that results are collected, print the total
# print("Linear feet of ALL pipes in the model is: {}".format(total_length))