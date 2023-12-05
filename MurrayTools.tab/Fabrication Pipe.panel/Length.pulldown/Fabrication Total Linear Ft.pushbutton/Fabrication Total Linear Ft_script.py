__title__ = 'Total\nLength'
__doc__ = """Calculates Combined Length of ALL Fabrication Pipes in the model."""

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter

doc = __revit__.ActiveUIDocument.Document
curview = doc.ActiveView

# Creating a collector instance and collecting all the Fabrication Pipe and Fittings
Pipe_collector = FilteredElementCollector(doc, curview.Id) \
                    .OfCategory(BuiltInCategory.OST_FabricationPipework) \
                    .WhereElementIsNotElementType()


# Iterate over pipes and collect Length data

total_length = 0.0

for pipe in Pipe_collector:
    if pipe.IsAStraight:
        len_param = pipe.Parameter[BuiltInParameter.FABRICATION_PART_LENGTH]
        if len_param:
            total_length = total_length + len_param.AsDouble()

# now that results are collected, print the total
print("Linear feet of ALL pipes in the model is: {}".format(total_length))
