import Autodesk
from Autodesk.Revit.DB import BoundingBoxXYZ, XYZ, Transaction, Transform, ViewType, BuiltInParameter, UnitTypeId, UnitUtils, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI.Selection import PickBoxStyle
from Autodesk.Revit.UI import TaskDialog
# Add Windows Forms references
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System.Windows.Forms import Form, Label, Button, DialogResult, FormBorderStyle, FormStartPosition, Control, AnchorStyles, FlowLayoutPanel, TextBox, CheckBox
from System import Array
from System.Drawing import Point, Size, Color, Font

# Define active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
rvt_year = int(app.VersionNumber)

# Unit conversion function
def convert_internal_units(value, get_internal=True, units='m'):
    if units == 'm':
        units = UnitTypeId.Meters
    if get_internal:
        return UnitUtils.ConvertToInternalUnits(value, units)
    return UnitUtils.ConvertFromInternalUnits(value, units)

# Custom WinForms dialog for view selection
class ViewSelectionForm(Form):
    def __init__(self, view_list):
        self.selected_views = []
        self.view_list = sorted(view_list, key=lambda v: v.Name)
        self.padding = 10
        self.button_spacing = 10
        self.scale_factor = self.get_dpi_scale()
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()

    def get_dpi_scale(self):
        screen = System.Windows.Forms.Screen.PrimaryScreen
        graphics = self.CreateGraphics()
        dpi_x = graphics.DpiX
        graphics.Dispose()
        return dpi_x / 96.0

    def scale_value(self, value):
        return int(value * self.scale_factor)

    def InitializeComponents(self):
        self.Text = "View Selection"
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MaximizeBox = True
        self.MinimizeBox = True
        self.StartPosition = FormStartPosition.CenterScreen

        # Instruction label
        self.label = Label()
        self.label.Text = "Search and select views:"
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
        self.checkbox_panel.AutoScroll = True
        self.checkbox_panel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        self.checkbox_panel.WrapContents = False
        self.checkbox_panel.Font = Font("Arial", self.scale_value(8))
        self.checkbox_panel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom

        # Add checkboxes for each view
        self.update_checkboxes(self.view_list)

        # OK Button
        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.AutoSize = True
        self.ok_button.Click += self.ok_clicked
        self.ok_button.Font = Font("Arial", self.scale_value(8))
        self.ok_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left

        # Cancel Button
        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.AutoSize = True
        self.cancel_button.Click += self.cancel_clicked
        self.cancel_button.Font = Font("Arial", self.scale_value(8))
        self.cancel_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left

        # Check All Button
        self.check_all_button = Button()
        self.check_all_button.Text = "Check All"
        self.check_all_button.AutoSize = True
        self.check_all_button.Click += self.check_all_clicked
        self.check_all_button.Font = Font("Arial", self.scale_value(8))
        self.check_all_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self.Controls.AddRange(Array[Control]([
            self.label, self.search_box, self.checkbox_panel,
            self.ok_button, self.cancel_button, self.check_all_button
        ]))

        # Calculate minimum width based on buttons
        total_button_width = (
            self.ok_button.Width +
            self.cancel_button.Width +
            self.check_all_button.Width +
            self.scale_value(2 * self.button_spacing) +
            self.scale_value(2 * self.padding)
        )
        min_width = max(self.scale_value(400), total_button_width)
        self.Size = Size(min_width, self.scale_value(400))
        self.MinimumSize = Size(min_width, self.scale_value(400))

        self.Resize += self.on_resize
        self.UpdateLayout()

    def update_checkboxes(self, views):
        self.checkbox_panel.Controls.Clear()
        self.checkboxes = []
        for view in views:
            checkbox = CheckBox()
            checkbox.Text = view.Name
            checkbox.Tag = view
            checkbox.AutoSize = True
            checkbox.Click += self.checkbox_clicked
            if view in self.selected_views:
                checkbox.Checked = True
            self.checkbox_panel.Controls.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered_views = [v for v in self.view_list if search_text in v.Name.lower()]
        self.update_checkboxes(filtered_views)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for checkbox in self.checkboxes:
            checkbox.Checked = self.check_all_state
        self.selected_views = [cb.Tag for cb in self.checkboxes if cb.Checked]

    def UpdateLayout(self):
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
        self.ok_button.Location = Point(start_x, button_y)
        self.cancel_button.Location = Point(start_x + self.ok_button.Width + self.scale_value(self.button_spacing), button_y)
        self.check_all_button.Location = Point(
            start_x + self.ok_button.Width + self.cancel_button.Width + self.scale_value(2 * self.button_spacing),
            button_y
        )

    def on_resize(self, sender, args):
        self.UpdateLayout()

    def checkbox_clicked(self, sender, args):
        self.selected_views = [cb.Tag for cb in self.checkboxes if cb.Checked]

    def ok_clicked(self, sender, args):
        self.selected_views = [cb.Tag for cb in self.checkboxes if cb.Checked]
        self.DialogResult = System.Windows.Forms.DialogResult.OK
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = System.Windows.Forms.DialogResult.Cancel
        self.Close()

# Collect all views (Floor Plans and Ceiling Plans)
view_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views) \
    .WhereElementIsNotElementType() \
    .ToElements()
all_views = [v for v in view_collector if v.ViewType in [ViewType.FloorPlan, ViewType.CeilingPlan]]

# Get pre-selected views
selected_view_ids = uidoc.Selection.GetElementIds()
selected_views = [doc.GetElement(vid) for vid in selected_view_ids if doc.GetElement(vid).ViewType in [ViewType.FloorPlan, ViewType.CeilingPlan]]

# If no views are pre-selected, show dialog
if not selected_views:
    if not all_views:
        TaskDialog.Show("Error", "No floor plans or ceiling plans found.")
        raise Exception("Script exited: No views found.")
    form = ViewSelectionForm(all_views)
    result = form.ShowDialog()
    if result != System.Windows.Forms.DialogResult.OK or not form.selected_views:
        TaskDialog.Show("Error", "No views selected. Please try again.")
        pass
        # raise Exception("Script exited: No views selected.")
    selected_views = form.selected_views

# Prompt user to select a box area
pickedBox = uidoc.Selection.PickBox(PickBoxStyle.Directional, "Select area for sketch")
if not pickedBox:
    TaskDialog.Show("Error", "No area selected. Please try again.")
    raise Exception("Script exited: No area selected.")

# Get max and min points of the selected box (in view coordinates)
newmax = pickedBox.Max
newmin = pickedBox.Min

# Check if view is set to True North
def is_true_north(view):
    if view.ViewType in [ViewType.FloorPlan, ViewType.CeilingPlan]:
        orient_param = view.get_Parameter(BuiltInParameter.PLAN_VIEW_NORTH)
        return orient_param and orient_param.AsInteger() == 1  # 1 = True North
    return False

# Adjust crop box based on view orientation
def adjust_crop_box(view, max_pt, min_pt):
    # Initialize min and max points
    adj_max = XYZ(max_pt.X, max_pt.Y, 0)
    adj_min = XYZ(min_pt.X, min_pt.Y, 0)
    
    if is_true_north(view) and view.CropBox.Transform:
        # Transform view coordinates to model space using view's transform
        transform = view.CropBox.Transform.Inverse
        adj_max = transform.OfPoint(adj_max)
        adj_min = transform.OfPoint(adj_min)
    
    # Ensure min/max ordering
    corrected_min = XYZ(min(adj_min.X, adj_max.X), min(adj_min.Y, adj_max.Y), 0)
    corrected_max = XYZ(max(adj_min.X, adj_max.X), max(adj_min.Y, adj_max.Y), 0)
    
    # Create adjusted bounding box
    adjusted_bbox = BoundingBoxXYZ()
    adjusted_bbox.Min = corrected_min
    adjusted_bbox.Max = corrected_max
    adjusted_bbox.Transform = Transform.Identity
    
    return adjusted_bbox

# Start transaction to apply changes
t = Transaction(doc, 'Set Crop Box for Selected Views')
t.Start()

for view in selected_views:
    adjusted_bbox = adjust_crop_box(view, newmax, newmin)
    view.CropBox = adjusted_bbox
    view.CropBoxActive = True
    view.CropBoxVisible = True

t.Commit()