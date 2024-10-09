from Autodesk.Revit.DB import BoundingBoxXYZ, FilteredElementCollector, Transaction, BuiltInCategory, FabricationPart, Level, LogicalOrFilter, ElementCategoryFilter
from rpw.ui.forms import FlexForm, Label, ComboBox, Button
import math

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Create a collector to get all Level elements in the document
level_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType()

# Store level names and ElementIds in a dictionary
level_elevations = {"(None)": None}  # Add (None) option with None as value
level_ids = {}  # Dictionary to store level names and their ElementId
for level in level_collector:
    level_elevations[level.Name] = level.Elevation
    level_ids[level.Name] = level.Id  # Store the ElementId of each level

# Sort levels by elevation to find the lowest and the next level
sorted_levels = sorted(level_elevations.items(), key=lambda x: x[1])  # Sort by elevation

# Pre-set the lowest level as Bottom and the next level up as Top
if len(sorted_levels) > 1:
    pre_set_bottom_level = sorted_levels[0][0]  # Lowest level
    pre_set_top_level = sorted_levels[1][0]     # Level above the lowest
else:
    pre_set_bottom_level = sorted_levels[0][0] if sorted_levels else "(None)"
    pre_set_top_level = "(None)"

# Function to get the center point of a fabrication part's bounding box
def GetCenterPoint(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    center = (bBox.Max + bBox.Min) / 2
    return center.Z  # Only return the Z (elevation) coordinate
try:
    # Display dialog to choose top and bottom levels, including a (None) option
    components = [
        Label('Choose Top Level:'),
        ComboBox('TopLevel', level_elevations.keys(), default=pre_set_top_level, sort=True),
        Label('Choose Bottom Level:'),
        ComboBox('BotLevel', level_elevations.keys(), default=pre_set_bottom_level, sort=True),
        Label('Parts are assigned to Bottom when both levels are set.'),
        Label('Parts are assigned to Bottom if Top is (None) ROOF.'),
        Label('Parts are assigned to Top if Bottom is (None) UG.'),
        Button('Ok')
    ]
    form = FlexForm('Select Levels', components)
    form.show()

    # Convert dialog input into variables for top and bottom levels
    TopLevelName = form.values['TopLevel']
    BotLevelName = form.values['BotLevel']

    # Get the Z elevations of the selected levels (None if "(None)" is selected)
    TopLevelElev = level_elevations[TopLevelName]
    BotLevelElev = level_elevations[BotLevelName]

    # Get the ElementIds of the selected levels (None if "(None)" is selected)
    TopLevelId = level_ids.get(TopLevelName, None)
    BotLevelId = level_ids.get(BotLevelName, None)

    # Create a category filter for Fabrication Pipework and Fabrication Ductwork
    pipework_filter = ElementCategoryFilter(BuiltInCategory.OST_FabricationPipework)
    ductwork_filter = ElementCategoryFilter(BuiltInCategory.OST_FabricationDuctwork)

    # Combine the filters with LogicalOrFilter
    combined_filter = LogicalOrFilter(pipework_filter, ductwork_filter)

    # Create a FilteredElementCollector to get all MEP Fabrication Pipework and Ductwork elements in the current view
    fabrication_elements = FilteredElementCollector(doc, curview.Id) \
                           .OfClass(FabricationPart) \
                           .WherePasses(combined_filter) \
                           .WhereElementIsNotElementType() \
                           .ToElements()

    # Start a transaction to modify the document
    t = Transaction(doc, "Assign Levels to Fabrication Parts")
    t.Start()

    # Iterate over the fabrication parts and check their center Z elevation
    for elem in fabrication_elements:
        try:
            center_z = GetCenterPoint(elem.Id)  # Get the Z center point of the fabrication part

            # If both levels are selected (not None), check if the center Z is between the levels
            if BotLevelElev is not None and TopLevelElev is not None:
                if BotLevelElev <= center_z <= TopLevelElev:
                    elem.LookupParameter("Reference Level").Set(BotLevelId)
                    print("Assigned {} to element {}".format(BotLevelName, elem.Id))

            # If Top Level is None, assign Bottom Level to all elements above the Bottom Level
            elif TopLevelElev is None and BotLevelElev is not None:
                if center_z >= BotLevelElev:
                    elem.LookupParameter("Reference Level").Set(BotLevelId)
                    print("Assigned {} to element {}".format(BotLevelName, elem.Id))

            # If Bottom Level is None, assign Top Level to all elements below the Top Level
            elif BotLevelElev is None and TopLevelElev is not None:
                if center_z <= TopLevelElev:
                    elem.LookupParameter("Reference Level").Set(TopLevelId)
                    print("Assigned {} to element {}".format(TopLevelName, elem.Id))

        except:
            pass
    # Commit the transaction after processing all elements
    t.Commit()
except:
    pass