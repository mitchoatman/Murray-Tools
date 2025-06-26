import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from Autodesk.Revit.DB import BoundingBoxXYZ, FilteredElementCollector, Transaction, BuiltInCategory, FabricationPart, Level, LogicalOrFilter, ElementCategoryFilter
from System.Windows.Forms import Form, Label, ComboBox, Button, FormStartPosition
from System.Drawing import Point, Size
from System import Array
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
sorted_levels = sorted(level_elevations.items(), key=lambda x: x[1] if x[1] is not None else float('inf'))  # Sort by elevation, handle None

# Pre-set the lowest level as Bottom and the next level up as Top
if len(sorted_levels) > 1:
    pre_set_bottom_level = sorted_levels[0][0]  # Lowest level
    pre_set_top_level = sorted_levels[1][0]     # Level above the lowest
else:
    pre_set_bottom_level = sorted_levels[0][0] if sorted_levels else "(None)"
    pre_set_top_level = "(None)"

# Function to get the bottom point Z elevation of a fabrication part's bounding box
def get_bottom_point_z(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    if bBox is None:
        return None
    return bBox.Min.Z  # Return the Z coordinate of the bottom point

# Create WinForms dialog
class LevelSelectionForm(Form):
    def __init__(self, levels, default_top, default_bottom):
        self.Text = "Select Levels"
        self.Size = Size(400, 300)
        self.StartPosition = FormStartPosition.CenterScreen
        self.levels = levels

        # Top Level Label
        self.top_label = Label()
        self.top_label.Text = "Choose Top Level:"
        self.top_label.Location = Point(20, 20)
        self.top_label.Size = Size(150, 20)
        self.Controls.Add(self.top_label)

        # Top Level ComboBox
        self.top_combo = ComboBox()
        self.top_combo.Location = Point(180, 20)
        self.top_combo.Size = Size(150, 20)
        self.top_combo.Items.AddRange(Array[object](sorted(levels.keys())))
        self.top_combo.SelectedItem = default_top
        self.Controls.Add(self.top_combo)

        # Bottom Level Label
        self.bottom_label = Label()
        self.bottom_label.Text = "Choose Bottom Level:"
        self.bottom_label.Location = Point(20, 60)
        self.bottom_label.Size = Size(150, 20)
        self.Controls.Add(self.bottom_label)

        # Bottom Level ComboBox
        self.bottom_combo = ComboBox()
        self.bottom_combo.Location = Point(180, 60)
        self.bottom_combo.Size = Size(150, 20)
        self.bottom_combo.Items.AddRange(Array[object](sorted(levels.keys())))
        self.bottom_combo.SelectedItem = default_bottom
        self.Controls.Add(self.bottom_combo)

        # Instruction Labels
        self.inst1 = Label()
        self.inst1.Text = "Parts are assigned to Bottom when both levels are set."
        self.inst1.Location = Point(20, 100)
        self.inst1.Size = Size(350, 20)
        self.Controls.Add(self.inst1)

        self.inst2 = Label()
        self.inst2.Text = "Parts are assigned to Bottom if Top is (None) ROOF."
        self.inst2.Location = Point(20, 120)
        self.inst2.Size = Size(350, 20)
        self.Controls.Add(self.inst2)

        self.inst3 = Label()
        self.inst3.Text = "Parts are assigned to Top if Bottom is (None) UG."
        self.inst3.Location = Point(20, 140)
        self.inst3.Size = Size(350, 20)
        self.Controls.Add(self.inst3)

        # OK Button
        self.ok_button = Button()
        self.ok_button.Text = "Ok"
        self.ok_button.Location = Point(150, 200)
        self.ok_button.Click += self.on_ok
        self.Controls.Add(self.ok_button)

        self.result = None

    def on_ok(self, sender, event):
        self.result = {
            "TopLevel": self.top_combo.SelectedItem,
            "BotLevel": self.bottom_combo.SelectedItem
        }
        self.DialogResult = System.Windows.Forms.DialogResult.OK
        self.Close()

try:
    # Display WinForms dialog
    form = LevelSelectionForm(level_elevations, pre_set_top_level, pre_set_bottom_level)
    if form.ShowDialog() == System.Windows.Forms.DialogResult.OK and form.result:
        # Convert dialog input into variables for top and bottom levels
        TopLevelName = form.result["TopLevel"]
        BotLevelName = form.result["BotLevel"]

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

        # Iterate over the fabrication parts and check their bottom Z elevation
        for elem in fabrication_elements:
            try:
                bottom_z = get_bottom_point_z(elem.Id)  # Get the Z bottom point of the fabrication part
                if bottom_z is None:
                    continue

                # If both levels are selected (not None), check if the bottom Z is between the levels
                if BotLevelElev is not None and TopLevelElev is not None:
                    if BotLevelElev <= bottom_z <= TopLevelElev:
                        elem.LookupParameter("Reference Level").Set(BotLevelId)
                        # print("Assigned {} to element {}".format(BotLevelName, elem.Id))

                # If Top Level is None, assign Bottom Level to all elements above the Bottom Level
                elif TopLevelElev is None and BotLevelElev is not None:
                    if bottom_z >= BotLevelElev:
                        elem.LookupParameter("Reference Level").Set(BotLevelId)
                        # print("Assigned {} to element {}".format(BotLevelName, elem.Id))

                # If Bottom Level is None, assign Top Level to all elements below the Top Level
                elif BotLevelElev is None and TopLevelElev is not None:
                    if bottom_z <= TopLevelElev:
                        elem.LookupParameter("Reference Level").Set(TopLevelId)
                        # print("Assigned {} to element {}".format(TopLevelName, elem.Id))

            except Exception as e:
                print("Error processing element {}: {}".format(elem.Id, str(e)))
        # Commit the transaction after processing all elements
        t.Commit()
except Exception as e:
    print("Script failed: {}".format(str(e)))