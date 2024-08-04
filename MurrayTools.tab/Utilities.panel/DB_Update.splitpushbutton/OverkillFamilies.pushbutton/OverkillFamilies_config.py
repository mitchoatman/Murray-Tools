import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, ElementCategoryFilter, FamilyInstance
from pyrevit import revit, DB, forms

#define the active Revit application and document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView


# List of BuiltInCategories considered as model categories
model_categories = {
    "Structural Framing": BuiltInCategory.OST_StructuralFraming,
    "Structural Columns": BuiltInCategory.OST_StructuralColumns,
    "Structural Stiffener": BuiltInCategory.OST_StructuralStiffener,
    "Pipe Accessory": BuiltInCategory.OST_PipeAccessory,
    "Plumbing Fixtures": BuiltInCategory.OST_PlumbingFixtures,
    "Mechanical Equipment": BuiltInCategory.OST_MechanicalEquipment,
    "Duct Accessory": BuiltInCategory.OST_DuctAccessory,
    "Generic Models": BuiltInCategory.OST_GenericModel,
    # Add more model categories as needed
}

# Create a filter for model categories
category_filters = [ElementCategoryFilter(bic) for bic in model_categories.values()]

# Create a collector with all model category filters combined
collector = FilteredElementCollector(doc).WherePasses(category_filters[0])

for filter in category_filters[1:]:
    collector = collector.UnionWith(FilteredElementCollector(doc).WherePasses(filter))

# Collect unique categories
all_categories = set()

for element in collector:
    category = element.Category
    if category is not None:
        all_categories.add(category.Name)

try:
    GroupOptions = sorted(model_categories.keys())

    selected_category = forms.SelectFromList.show(GroupOptions, group_selector_title='Category:', multiselect=True, button_name='Select Category', exitscript=True, title="Overkill")

    if selected_category:
        for cat in selected_category:
            res = model_categories[cat]

            def GetCenterPoint(ele):
                bBox = doc.GetElement(ele).get_BoundingBox(None)
                center = (bBox.Max + bBox.Min) / 2
                return (center.X, center.Y, center.Z)

            def IsNestedFamily(element):
                # Check if the element is a FamilyInstance and if it has a parent (indicating it's nested)
                if isinstance(element, FamilyInstance):
                    return element.SuperComponent is not None
                return False

            # Create a FilteredElementCollector to get all elements of the selected category
            AllElements = FilteredElementCollector(doc).OfCategory(res).WhereElementIsNotElementType().ToElements()

            # Filter out nested families
            main_families = [el for el in AllElements if not IsNestedFamily(el)]

            # Get the center point of each selected element
            element_ids = []
            center_points = []

            for reference in main_families:
                center_point = GetCenterPoint(reference.Id)
                center_points.append(center_point)
                element_ids.append(reference.Id)

            # Find the duplicates in the list of center points
            duplicates = []
            duplicate_element_ids = []
            unique_center_points = []

            for i, cp in enumerate(center_points):
                if cp not in unique_center_points:
                    unique_center_points.append(cp)
                else:
                    duplicates.append(cp)
                    duplicate_element_ids.append(element_ids[i])

            try:
                if duplicates:
                    forms.alert_ifnot(len(duplicates) < 0,
                                      ("Delete Duplicate(s) in {}: {}".format(cat, len(duplicates))),
                                      yes=True, no=True, exitscript=True)
                    
                    with Transaction(doc, "Delete Elements") as transaction:
                        transaction.Start()
                        for element_id in duplicate_element_ids:
                            doc.Delete(element_id)
                        transaction.Commit()
                else:
                    forms.show_balloon('Duplicates', 'No Duplicates Found')

            except:
                pass
except:
    pass
