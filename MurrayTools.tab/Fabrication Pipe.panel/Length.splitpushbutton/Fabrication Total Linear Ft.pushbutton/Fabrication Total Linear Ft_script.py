from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FabricationPart

doc = __revit__.ActiveUIDocument.Document

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
    import sys
    sys.exit()

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()

# Dictionary to store total length for each level and material
level_material_total_lengths = {}
level_total_lengths = {}
material_total_lengths = {}
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

# Print the results
print("Linear feet of ALL pipes in the model: [ {} ]".format(total_length))
print("----------\n\nTotal Lengths of Fabrication Pipes by Level:")
for level, length in level_total_lengths.items():
    print("  {}: [ {} ] Linear feet".format(level, length))

print("----------\n\nTotal Lengths of Fabrication Pipes by Level and Material:")
for level, materials in level_material_total_lengths.items():
    print("\nLevel: {}".format(level))
    for material, length in materials.items():
        print("  Material: {}: [ {} ] Linear feet".format(material, length))

print("----------\n\nTotal Lengths of Fabrication Pipes by Material:")
for material, length in material_total_lengths.items():
    print("  {}: [ {} ] Linear feet".format(material, length))