import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, \
    ParameterValueProvider, ElementId, FilterStringBeginsWith, Transaction, FilterStringEquals, \
    FilterStringLessOrEqual, FilterStringRule, ElementParameterFilter, ParameterValueProvider, LogicalOrFilter, TemporaryViewMode
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString

# Add Windows Forms references
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System.Windows.Forms import (Form, Label, Button, DialogResult, 
                                 FormBorderStyle, FormStartPosition, Control, AnchorStyles, FlowLayoutPanel, TextBox)
from System import Array
from System.Drawing import Point, Size, Color, Font

doc = __revit__.ActiveUIDocument.Document
DB = Autodesk.Revit.DB
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

duct_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationDuctwork) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

def create_filter_2023_newer(key_parameter, element_value):
    """Function to create a filter from builtinParameter and Value."""
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    f_parameter_value = element_value
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), f_parameter_value)
    my_filter = ElementParameterFilter(f_rule)
    return my_filter

def create_filter_2022_older(key_parameter, element_value):
    """Function to create a filter from builtinParameter and Value."""
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    f_parameter_value = element_value
    caseSensitive = False
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), f_parameter_value, caseSensitive)
    my_filter = ElementParameterFilter(f_rule)
    return my_filter

# Define the WinForms dialog class with DPI scaling, auto-size buttons, checkboxes, search bar, and toggle check all
class ServiceSelectionForm(Form):
    def __init__(self, service_list):
        self.selected_services = []
        self.service_list = sorted(service_list)
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
        self.Text = "Service Visibility"
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MaximizeBox = True
        self.MinimizeBox = True
        self.StartPosition = FormStartPosition.CenterScreen

        # Instruction label
        self.label = Label()
        self.label.Text = "Search and select services:"
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

        # Add checkboxes for each service
        self.update_checkboxes(self.service_list)

        # Reset Button
        self.reset_button = Button()
        self.reset_button.Text = "Reset View"
        self.reset_button.AutoSize = True
        self.reset_button.BackColor = Color.FromArgb(128, 255, 0, 0)
        self.reset_button.Click += self.reset_clicked
        self.reset_button.Font = Font("Arial", self.scale_value(8))
        self.reset_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left

        # Hide Button
        self.hide_button = Button()
        self.hide_button.Text = "Hide"
        self.hide_button.AutoSize = True
        self.hide_button.Click += self.hide_clicked
        self.hide_button.Font = Font("Arial", self.scale_value(8))
        self.hide_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left

        # Isolate Button
        self.isolate_button = Button()
        self.isolate_button.Text = "Isolate"
        self.isolate_button.AutoSize = True
        self.isolate_button.Click += self.isolate_clicked
        self.isolate_button.Font = Font("Arial", self.scale_value(8))
        self.isolate_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left

        # Check All Button (on the right with same spacing)
        self.check_all_button = Button()
        self.check_all_button.Text = "Check All"
        self.check_all_button.AutoSize = True  # Auto-size to fit text
        self.check_all_button.Click += self.check_all_clicked
        self.check_all_button.Font = Font("Arial", self.scale_value(8))
        self.check_all_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self.Controls.AddRange(Array[Control]([
            self.label, self.search_box, self.checkbox_panel, self.reset_button, 
            self.hide_button, self.isolate_button, self.check_all_button
        ]))

        # Calculate minimum width based on buttons
        total_button_width = (
            self.reset_button.Width + 
            self.hide_button.Width + 
            self.isolate_button.Width + 
            self.check_all_button.Width + 
            self.scale_value(3 * self.button_spacing) +  # Three gaps between four buttons
            self.scale_value(2 * self.padding)  # 10pt padding on left and right
        )
        min_width = max(self.scale_value(400), total_button_width)
        self.Size = Size(min_width, self.scale_value(400))
        self.MinimumSize = Size(min_width, self.scale_value(400))

        # Handle resize event
        self.Resize += self.on_resize
        self.UpdateLayout()

    def update_checkboxes(self, services):
        """Update the checkbox panel with filtered services."""
        self.checkbox_panel.Controls.Clear()
        self.checkboxes = []
        for service in services:
            checkbox = System.Windows.Forms.CheckBox()
            checkbox.Text = service
            checkbox.AutoSize = True
            checkbox.Click += self.checkbox_clicked
            # Restore checked state if previously selected
            if service in self.selected_services:
                checkbox.Checked = True
            self.checkbox_panel.Controls.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        """Filter checkboxes based on search input."""
        search_text = self.search_box.Text.lower()
        filtered_services = [s for s in self.service_list if search_text in s.lower()]
        self.update_checkboxes(filtered_services)

    def check_all_clicked(self, sender, args):
        """Toggle between checking all and unchecking all visible checkboxes."""
        self.check_all_state = not self.check_all_state  # Toggle state
        for checkbox in self.checkboxes:
            checkbox.Checked = self.check_all_state
        self.selected_services = [cb.Text for cb in self.checkboxes if cb.Checked]

    def UpdateLayout(self):
        # Update search box width
        self.search_box.Size = Size(
            self.ClientSize.Width - self.scale_value(2 * self.padding),
            self.scale_value(20)
        )
        # Update checkbox panel size
        self.checkbox_panel.Size = Size(
            self.ClientSize.Width - self.scale_value(2 * self.padding),
            self.ClientSize.Height - self.scale_value(self.padding + 40 + self.padding + 23 + self.padding)
        )
        # Position buttons dynamically
        start_x = self.scale_value(self.padding)
        button_y = self.ClientSize.Height - self.scale_value(self.padding + 23)
        
        # Left-aligned buttons
        self.reset_button.Location = Point(start_x, button_y)
        self.hide_button.Location = Point(start_x + self.reset_button.Width + self.scale_value(self.button_spacing), button_y)
        self.isolate_button.Location = Point(start_x + self.reset_button.Width + self.hide_button.Width + self.scale_value(2 * self.button_spacing), button_y)
        # Right-aligned Check All button with consistent spacing
        self.check_all_button.Location = Point(
            start_x + self.reset_button.Width + self.hide_button.Width + self.isolate_button.Width + self.scale_value(3 * self.button_spacing),
            button_y
        )

    def on_resize(self, sender, args):
        self.UpdateLayout()

    def checkbox_clicked(self, sender, args):
        # Update selected_services based on checked state
        self.selected_services = [cb.Text for cb in self.checkboxes if cb.Checked]

    def reset_clicked(self, sender, args):
        try:
            t = Transaction(doc, "Reset Temporary Hide/Isolate")
            t.Start()
            curview.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
            t.Commit()
            self.Close()
        except Exception as e:
            print("Reset Error: {}".format(str(e)))

    def hide_clicked(self, sender, args):
        self.selected_services = [cb.Text for cb in self.checkboxes if cb.Checked]
        if self.selected_services:
            try:
                list_of_filters = list()
                if RevitINT > 2022:
                    for fp_servicename in self.selected_services:
                        cat_filter = create_filter_2023_newer(
                            key_parameter=BuiltInParameter.FABRICATION_SERVICE_NAME, 
                            element_value=str(fp_servicename)
                        )
                        list_of_filters.append(cat_filter)
                else:
                    for fp_servicename in self.selected_services:
                        cat_filter = create_filter_2022_older(
                            key_parameter=BuiltInParameter.FABRICATION_SERVICE_NAME, 
                            element_value=str(fp_servicename)
                        )
                        list_of_filters.append(cat_filter)

                if list_of_filters:
                    multiple_filters = LogicalOrFilter(list_of_filters)
                    analyticalCollector = FilteredElementCollector(doc).WherePasses(multiple_filters).ToElementIds()
                    
                    t = Transaction(doc, "Hide Services")
                    t.Start()
                    curview.HideElementsTemporary(analyticalCollector)
                    t.Commit()
                    self.Close()
            except Exception as e:
                print("Hide Error: {}".format(str(e)))

    def isolate_clicked(self, sender, args):
        self.selected_services = [cb.Text for cb in self.checkboxes if cb.Checked]
        if self.selected_services:
            try:
                list_of_filters = list()
                if RevitINT > 2022:
                    for fp_servicename in self.selected_services:
                        cat_filter = create_filter_2023_newer(
                            key_parameter=BuiltInParameter.FABRICATION_SERVICE_NAME, 
                            element_value=str(fp_servicename)
                        )
                        list_of_filters.append(cat_filter)
                else:
                    for fp_servicename in self.selected_services:
                        cat_filter = create_filter_2022_older(
                            key_parameter=BuiltInParameter.FABRICATION_SERVICE_NAME, 
                            element_value=str(fp_servicename)
                        )
                        list_of_filters.append(cat_filter)

                if list_of_filters:
                    multiple_filters = LogicalOrFilter(list_of_filters)
                    analyticalCollector = FilteredElementCollector(doc).WherePasses(multiple_filters).ToElementIds()
                    
                    t = Transaction(doc, "Isolate Services")
                    t.Start()
                    curview.IsolateElementsTemporary(analyticalCollector)
                    t.Commit()
                    self.Close()
            except Exception as e:
                print("Isolate Error: {}".format(str(e)))

# Collect services
SrvcList = list()
for Item in hanger_collector:
    servicename = get_parameter_value_by_name_AsString(Item, 'Fabrication Service Name')
    SrvcList.append(servicename)

for Item in pipe_collector:
    servicename = get_parameter_value_by_name_AsString(Item, 'Fabrication Service Name')
    SrvcList.append(servicename)

for Item in duct_collector:
    servicename = get_parameter_value_by_name_AsString(Item, 'Fabrication Service Name')
    SrvcList.append(servicename)

# Remove duplicates and show form
unique_services = set(SrvcList)
if not unique_services:
    print("No fabrication services found in the current view.")
    import sys
    sys.exit()

form = ServiceSelectionForm(unique_services)
form.ShowDialog()