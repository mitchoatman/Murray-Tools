import Autodesk.Revit.DB as DB

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

def gradient(grid):
    start = grid.Curve.GetEndPoint(0)
    end = grid.Curve.GetEndPoint(1)
    return round((start.Y - end.Y) / (start.X - end.X), 10) if round(start.X, 10) != round(end.X, 10) else None

def refArray(grids):
    ref_array = DB.ReferenceArray()
    for grid in grids:
        ref_array.Append(DB.Reference(grid))
    return ref_array

def refLine(grids):
    return DB.Line.CreateBound(grids[0].Curve.GetEndPoint(0), grids[1].Curve.GetEndPoint(0))

# Collect grids
grids = list(DB.FilteredElementCollector(doc)
             .WherePasses(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Grids))
             .WhereElementIsNotElementType())

# Group parallel grids
grid_groups = {}
excluded = set()
for grid in grids:
    grid_name = grid.LookupParameter("Name").AsString()
    if grid_name not in excluded:
        group = [grid]
        grid_curve = grid.Curve
        grid_grad = gradient(grid)
        excluded.add(grid_name)
        
        for other in grids:
            other_name = other.LookupParameter("Name").AsString()
            if (other_name not in excluded and 
                grid_curve.Intersect(other.Curve) == DB.SetComparisonResult.Disjoint and 
                grid_grad == gradient(other)):
                group.append(other)
                excluded.add(other_name)
        
        if len(group) > 1:
            grid_groups[grid_name] = group

# Create dimensions
with DB.Transaction(doc, "Dimension grids") as t:
    t.Start()
    for grids in grid_groups.values():
        try:
            doc.Create.NewDimension(view, refLine(grids), refArray(grids))
        except:
            pass
    t.Commit()