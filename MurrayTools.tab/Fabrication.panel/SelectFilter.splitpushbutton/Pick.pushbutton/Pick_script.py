from collections import namedtuple

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import System
from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, ElementId
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

# Add Windows Forms references
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System.Windows.Forms import (Form, Label, Button, FormBorderStyle, 
                                 FormStartPosition, Control, AnchorStyles, 
                                 FlowLayoutPanel, TextBox)
from System import Array
from System.Drawing import Point, Size, Font

# Categories to exclude
CAT_EXCLUDED = (
    int(DB.BuiltInCategory.OST_RoomSeparationLines),
    int(DB.BuiltInCategory.OST_Cameras),
    int(DB.BuiltInCategory.OST_CurtainGrids),
    int(DB.BuiltInCategory.OST_Elev),
    int(DB.BuiltInCategory.OST_IOSModelGroups),
    int(DB.BuiltInCategory.OST_SitePropertyLineSegment),
    int(DB.BuiltInCategory.OST_SectionBox),
    int(DB.BuiltInCategory.OST_ShaftOpening),
    int(DB.BuiltInCategory.OST_BeamAnalytical),
    int(DB.BuiltInCategory.OST_StructuralFramingOpening),
    int(DB.BuiltInCategory.OST_MEPSpaceSeparationLines),
    int(DB.BuiltInCategory.OST_DuctSystem),
    int(DB.BuiltInCategory.OST_PipingSystem),
    int(DB.BuiltInCategory.OST_CenterLines),
    int(DB.BuiltInCategory.OST_PipeCurvesCenterLine),
    int(DB.BuiltInCategory.OST_FabricationPipeworkCenterLine),
    int(DB.BuiltInCategory.OST_FabricationPipeworkDrop),
    int(DB.BuiltInCategory.OST_FabricationPipeworkRise),
    int(DB.BuiltInCategory.OST_FabricationPipeworkInsulation),
    int(DB.BuiltInCategory.OST_CurtainGridsRoof),
    int(DB.BuiltInCategory.OST_SWallRectOpening),
    -2000278,
    -1,
)

CategoryOption = namedtuple('CategoryOption', ['name', 'revit_cat'])

class PickByCategorySelectionFilter(ISelectionFilter):
    """Selection filter implementation for multiple categories"""
    def __init__(self, category_opts):
        self.category_ids = [cat_opt.revit_cat.Id for cat_opt in category_opts]

    def AllowElement(self, element):
        """Is element allowed to be selected?"""
        if element.Category and element.Category.Id in self.category_ids:
            return True
        return False

    def AllowReference(self, refer, point):  # pylint: disable=W0613
        """Not used for selection"""
        return False

def get_categories_in_active_view(doc, view):
    """Get all unique categories present in the active view, excluding specified categories"""
    collector = FilteredElementCollector(doc, view.Id)
    elements = collector.WhereElementIsNotElementType().ToElements()
    # Use set to ensure unique category IDs
    category_ids = set()
    category_dict = {}
    for element in elements:
        if (element.Category and element.Category.Name and 
            element.Category.Id.IntegerValue not in CAT_EXCLUDED and 
            element.Category.Id not in category_ids):
            category_ids.add(element.Category.Id)
            category_dict[element.Category.Id] = element.Category
    # Sort categories by name for consistent dialog display
    return sorted(category_dict.values(), key=lambda x: x.Name)

def pick_by_categories(category_opts, uidoc):
    """Handle selection by category"""
    msfilter = PickByCategorySelectionFilter(category_opts)
    selection_list = uidoc.Selection.PickElementsByRectangle(msfilter)

    if selection_list:
        # Extract Ids properly
        filtered_ids = List[ElementId]([e.Id for e in selection_list])
        uidoc.Selection.SetElementIds(filtered_ids)

# Define the WinForms dialog class with DPI scaling, checkboxes, search bar, and check all button
class FilterSelectionByCategory(Form):
    def __init__(self, category_list):
        self.selected_categories = []
        self.category_list = sorted(category_list, key=lambda x: x.Name)
        self.padding = 10  # Base padding constant
        self.button_spacing = 10  # Base space between buttons
        self.scale_factor = self.get_dpi_scale()  # Get DPI scaling factor
        self.checkboxes = []  # To store checkbox controls
        self.check_all_state = False  # Track toggle state
        self.InitializeComponents()

    def get_dpi_scale(self):
        """Calculate the DPI scaling factor based on the primary screen."""
        screen = System.Windows.Forms.Screen.PrimaryScreen
        graphics = self.CreateGraphics()  # Create a graphics object to get DPI
        dpi_x = graphics.DpiX
        graphics.Dispose()  # Clean up
        return dpi_x / 96.0  # 96 DPI is the default (100%)

    def scale_value(self, value):
        """Scale a value based on the DPI scaling factor."""
        return int(value * self.scale_factor)

    def InitializeComponents(self):
        self.Text = "Filter Selection by Category"
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MaximizeBox = True
        self.MinimizeBox = True
        self.StartPosition = FormStartPosition.CenterScreen

        # Instruction label
        self.label = Label()
        self.label.Text = "Search and select categories:"
        self.label.Location = Point(self.scale_value(self.padding), self.scale_value(self.padding))
        self.label.AutoSize = True
        self.label.Font = Font("Arial", self.scale_value(8))
        self.label.Anchor = AnchorStyles.Top | AnchorStyles.Left

        # Search TextBox
        self.search_box = TextBox()
        self.search_box.Location = Point(self.scale_value(self.padding), self.scale_value(self.padding + 20))
        self.search_box.Font = Font("Arial", self.scale_value(8))
        self.search_box.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.search_box.TextChanged += self.search_changed

        # FlowLayoutPanel for checkboxes
        self.checkbox_panel = FlowLayoutPanel()
        self.checkbox_panel.Location = Point(self.scale_value(self.padding), self.scale_value(self.padding + 40))
        self.checkbox_panel.AutoScroll = True  # Enable scrolling if content overflows
        self.checkbox_panel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        self.checkbox_panel.WrapContents = False
        self.checkbox_panel.Font = Font("Arial", self.scale_value(8))
        self.checkbox_panel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom

        # Add checkboxes for each category
        self.update_checkboxes(self.category_list)

        # Select Button
        self.select_button = Button()
        self.select_button.Text = "Select"
        self.select_button.AutoSize = True
        self.select_button.Click += self.select_clicked
        self.select_button.Font = Font("Arial", self.scale_value(8))
        self.select_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left

        # Check All Button
        self.check_all_button = Button()
        self.check_all_button.Text = "Check All"
        self.check_all_button.AutoSize = True
        self.check_all_button.Click += self.check_all_clicked
        self.check_all_button.Font = Font("Arial", self.scale_value(8))
        self.check_all_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self.Controls.AddRange(Array[Control]([
            self.label, self.search_box, self.checkbox_panel, 
            self.select_button, self.check_all_button
        ]))

        # Calculate minimum width based on buttons
        total_button_width = (
            self.select_button.Width + 
            self.check_all_button.Width + 
            self.scale_value(self.button_spacing) +  # One gap between two buttons
            self.scale_value(2 * self.padding)  # Padding on left and right
        )
        min_width = max(self.scale_value(400), total_button_width)
        self.Size = Size(min_width, self.scale_value(400))
        self.MinimumSize = Size(min_width, self.scale_value(400))

        # Handle resize event
        self.Resize += self.on_resize
        self.UpdateLayout()

    def update_checkboxes(self, categories):
        """Update the checkbox panel with filtered categories."""
        self.checkbox_panel.Controls.Clear()
        self.checkboxes = []
        for category in categories:
            checkbox = System.Windows.Forms.CheckBox()
            checkbox.Text = category.Name
            checkbox.Tag = category  # Store the category object
            checkbox.AutoSize = True
            checkbox.Click += self.checkbox_clicked
            if category.Name in self.selected_categories:
                checkbox.Checked = True
            self.checkbox_panel.Controls.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        """Filter checkboxes based on search input."""
        search_text = self.search_box.Text.lower()
        filtered_categories = [c for c in self.category_list if search_text in c.Name.lower()]
        self.update_checkboxes(filtered_categories)

    def check_all_clicked(self, sender, args):
        """Toggle between checking all and unchecking all visible checkboxes."""
        self.check_all_state = not self.check_all_state
        for checkbox in self.checkboxes:
            checkbox.Checked = self.check_all_state
        self.selected_categories = [cb.Text for cb in self.checkboxes if cb.Checked]

    def checkbox_clicked(self, sender, args):
        """Update selected_categories based on checked state."""
        self.selected_categories = [cb.Text for cb in self.checkboxes if cb.Checked]

    def select_clicked(self, sender, args):
        """Handle select button click to store selected categories and close form."""
        self.selected_categories = [cb.Text for cb in self.checkboxes if cb.Checked]
        self.DialogResult = System.Windows.Forms.DialogResult.OK
        self.Close()

    def UpdateLayout(self):
        """Update control positions on form resize."""
        self.search_box.Size = Size(
            self.ClientSize.Width - self.scale_value(2 * self.padding),
            self.scale_value(20)
        )
        self.checkbox_panel.Size = Size(
            self.ClientSize.Width - self.scale_value(2 * self.padding),
            self.ClientSize.Height - self.scale_value(self.padding + 40 + self.padding + 23 + self.padding)
        )
        start_x = self.scale_value(self.padding)
        button_y = self.ClientSize.Height - self.scale_value(self.padding + 23)
        self.select_button.Location = Point(start_x, button_y)
        self.check_all_button.Location = Point(
            start_x + self.select_button.Width + self.scale_value(self.button_spacing),
            button_y
        )

    def on_resize(self, sender, args):
        self.UpdateLayout()

# Get Revit document and UI document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Get unique categories present in the active view
source_categories = get_categories_in_active_view(doc, doc.ActiveView)

# Show form and process selected categories
if source_categories:
    form = FilterSelectionByCategory(source_categories)
    if form.ShowDialog() == System.Windows.Forms.DialogResult.OK and form.selected_categories:
        category_opts = [CategoryOption(name=x.Name, revit_cat=x) 
                         for x in source_categories if x.Name in form.selected_categories]
        pick_by_categories(category_opts, uidoc)
else:
    print("No categories found in the current view.")