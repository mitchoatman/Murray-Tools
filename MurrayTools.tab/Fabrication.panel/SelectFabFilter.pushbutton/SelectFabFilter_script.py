# -*- coding: utf-8 -*-
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Form, Label, ComboBox, Button, ListBox, 
                                CheckBox, DialogResult, FormBorderStyle, FormStartPosition, 
                                SelectionMode, Control, ComboBoxStyle, TextBox, MessageBox)
from System import Array
from System.Drawing import Point, Size
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, FabricationConfiguration
from pyrevit import revit
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import (get_parameter_value_by_name_AsString, 
                                     get_parameter_value_by_name_AsValueString, 
                                     get_parameter_value_by_name_AsInteger)

Shared_Params()

# Define the active Revit document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

class RemoveFilterDialog(Form):
    def __init__(self, filter_options, filter_keys):
        self.filter_options = filter_options
        self.filter_keys = filter_keys
        self.selected_filter = None
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Text = "Remove Filter"
        self.Size = Size(300, 150)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterParent

        self.label = Label()
        self.label.Text = "Select filter to remove:"
        self.label.Location = Point(10, 10)
        self.label.Size = Size(280, 20)

        self.filter_combo = ComboBox()
        self.filter_combo.Location = Point(10, 30)
        self.filter_combo.Size = Size(260, 20)
        self.filter_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.filter_combo.Items.AddRange(Array[object](self.filter_options))
        if self.filter_options:
            self.filter_combo.SelectedIndex = 0

        self.ok_button = Button()
        self.ok_button.Text = "Remove"
        self.ok_button.Location = Point(110, 70)
        self.ok_button.AutoSize = True
        self.ok_button.Click += self.ok_clicked

        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Location = Point(190, 70)
        self.cancel_button.AutoSize = True
        self.cancel_button.Click += self.cancel_clicked

        self.Controls.AddRange(Array[Control]([
            self.label, self.filter_combo, self.ok_button, self.cancel_button
        ]))

    def ok_clicked(self, sender, args):
        if self.filter_combo.SelectedIndex >= 0:
            self.selected_filter = self.filter_keys[self.filter_combo.SelectedIndex]
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

class MultiPropertyFilterForm(Form):
    def __init__(self, property_options):
        self.property_options = property_options
        self.InitializeComponents()
        self.selected_filters = {}  # {property: [(values, is_and), ...]}
        if self.property_options:
            first_property = self.property_combo.Items[0]
            if first_property in self.property_options:
                self.property_combo.SelectedItem = first_property
            self.property_changed(None, None)
        self.update_filter_display()

    def InitializeComponents(self):
        self.Text = "Multi-Property Filter"
        self.Size = Size(550, 550)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        # Property selection
        self.property_label = Label()
        self.property_label.Text = "Select Property:"
        self.property_label.Location = Point(10, 10)
        self.property_label.Size = Size(120, 20)

        self.property_combo = ComboBox()
        self.property_combo.Location = Point(130, 10)
        self.property_combo.Size = Size(160, 20)
        self.property_combo.DropDownStyle = ComboBoxStyle.DropDownList
        properties = list(self.property_options.keys())
        if properties:
            self.property_combo.Items.AddRange(Array[object](properties))
        self.property_combo.SelectedIndexChanged += self.property_changed
        self.Controls.Add(self.property_combo)

        # AND/OR toggle (right of dropdown)
        self.logic_check = CheckBox()
        self.logic_check.Text = "AND logic (unchecked = OR)"
        self.logic_check.Location = Point(300, 10)
        self.logic_check.Size = Size(230, 20)

        # Search label
        self.search_label = Label()
        self.search_label.Text = "Search:"
        self.search_label.Location = Point(10, 40)
        self.search_label.Size = Size(120, 20)

        # Search bar
        self.search_box = TextBox()
        self.search_box.Location = Point(130, 40)
        self.search_box.Size = Size(300, 20)
        self.search_box.TextChanged += self.search_changed
        self.Controls.Add(self.search_box)

        # Values list
        self.values_label = Label()
        self.values_label.Text = "Select Values:"
        self.values_label.Location = Point(10, 70)
        self.values_label.Size = Size(120, 20)

        # Add Filter button (under "Select Values" label)
        self.add_button = Button()
        self.add_button.Text = "Add Filter"
        self.add_button.Location = Point(12, 90)
        self.add_button.AutoSize = True
        self.add_button.Click += self.add_filter

        # Remove Filter button (stacked below Add Filter)
        self.remove_button = Button()
        self.remove_button.Text = "Remove Filter"
        self.remove_button.Location = Point(12, 310)
        self.remove_button.AutoSize = True
        self.remove_button.Click += self.remove_filter

        self.values_list = ListBox()
        self.values_list.Location = Point(130, 70)
        self.values_list.Size = Size(350, 270)
        self.values_list.SelectionMode = SelectionMode.MultiExtended
        self.values_list.DoubleClick += self.add_filter

        # Filter feedback label - taller and full width
        self.filter_display = Label()
        self.filter_display.Text = "No filters added yet."
        self.filter_display.Location = Point(10, 350)
        self.filter_display.Size = Size(510, 100)
        self.filter_display.AutoSize = False

        # Buttons (centered at bottom, 15 pixels from bottom)
        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.Location = Point(130, 470)
        self.ok_button.AutoSize = True
        self.ok_button.Click += self.ok_clicked

        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Location = Point(350, 470)
        self.cancel_button.AutoSize = True
        self.cancel_button.Click += self.cancel_clicked

        self.Controls.AddRange(Array[Control]([
            self.property_label, 
            self.search_label,
            self.values_label, 
            self.values_list,
            self.logic_check, 
            self.add_button,
            self.remove_button,
            self.ok_button, 
            self.cancel_button,
            self.filter_display,
            self.search_box
        ]))

    def property_changed(self, sender, args):
        selected_property = self.property_combo.SelectedItem
        if selected_property:
            self.values_list.Items.Clear()
            self.search_box.Text = ""  # Clear search when property changes
            values = self.property_options.get(selected_property, [])
            if values:
                self.values_list.Items.AddRange(Array[object](values))

    def search_changed(self, sender, args):
        selected_property = self.property_combo.SelectedItem
        if selected_property:
            self.values_list.Items.Clear()
            search_term = self.search_box.Text.lower()
            values = self.property_options.get(selected_property, [])
            filtered_values = [v for v in values if search_term in str(v).lower()]
            if filtered_values:
                self.values_list.Items.AddRange(Array[object](filtered_values))

    def add_filter(self, sender, args):
        selected_property = self.property_combo.SelectedItem
        selected_values = [item for item in self.values_list.SelectedItems]
        if selected_property and selected_values:
            if selected_property not in self.selected_filters:
                self.selected_filters[selected_property] = []
            self.selected_filters[selected_property].append((selected_values, self.logic_check.Checked))
            self.update_filter_display()

    def remove_filter(self, sender, args):
        if not self.selected_filters:
            MessageBox.Show("No filters to remove.")
            return
        
        # Build a list of filter descriptions for the user to choose from
        filter_options = []
        filter_keys = []
        for prop, filter_list in self.selected_filters.items():
            for values, is_and in filter_list:
                mode = "AND" if is_and else "OR"
                desc = "%s (%s): %s" % (prop, mode, ", ".join(str(v) for v in values))
                filter_options.append(desc)
                filter_keys.append((prop, values))
        
        if not filter_options:
            MessageBox.Show("No filters to remove.")
            return
        
        # Show custom dialog to select filter
        dialog = RemoveFilterDialog(filter_options, filter_keys)
        if dialog.ShowDialog(self) == DialogResult.OK and dialog.selected_filter:
            prop, values = dialog.selected_filter
            for i, (v, is_and) in enumerate(self.selected_filters[prop]):
                if v == values:
                    del self.selected_filters[prop][i]
                    break
            # Clean up empty property entry
            if not self.selected_filters[prop]:
                del self.selected_filters[prop]
            self.update_filter_display()

    def update_filter_display(self, sender=None, args=None):
        if not self.selected_filters:
            self.filter_display.Text = "No filters added yet."
        else:
            filter_text = "Filters:\n"
            for prop, filter_list in self.selected_filters.items():
                filter_text += "%s:\n    " % prop  # Property name on its own line, conditions indented
                conditions = []
                for values, is_and in filter_list:
                    mode = "AND" if is_and else "OR"
                    condition = "[%s: %s]" % (mode, ", ".join(str(v) for v in values))
                    conditions.append(condition)
                filter_text += " ".join(conditions) + "\n"
            self.filter_display.Text = filter_text.strip()

    def ok_clicked(self, sender, args):
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

def get_property_value(elem, property_name):
    property_map = {
        'CID': lambda x: x.ItemCustomId,
        'ServiceType': lambda x: Config.GetServiceTypeName(x.ServiceType),
        'Name': lambda x: get_parameter_value_by_name_AsValueString(x, 'Family'),
        'Service Name': lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'),
        'Service Abbreviation': lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'),
        'Size': lambda x: get_parameter_value_by_name_AsString(x, 'Size of Primary End'),
        'STRATUS Assembly': lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Assembly'),
        'Line Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Line Number'),
        'STRATUS Status': lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Status'),
        'Reference Level': lambda x: get_parameter_value_by_name_AsValueString(x, 'Reference Level'),
        'Item Number': lambda x: get_parameter_value_by_name_AsString(x, 'Item Number'),
        'Bundle Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Bundle'),
        'REF BS Designation': lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF BS Designation'),
        'REF Line Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF Line Number'),
        'Comments': lambda x: get_parameter_value_by_name_AsString(x, 'Comments'),
        'Specification': lambda x: Config.GetSpecificationName(x.Specification),
        'Hanger Rod Size': lambda x: get_parameter_value_by_name_AsValueString(x, 'FP_Rod Size'),
        'Valve Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Valve Number')
    }
    try:
        return property_map.get(property_name, lambda x: None)(elem)
    except Exception, e:
        return None

# Collect elements and properties
preselection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
fab_elements = preselection if preselection else FilteredElementCollector(doc, curview.Id) \
    .OfClass(DB.FabricationPart) \
    .WhereElementIsNotElementType() \
    .ToElements()
all_elements = preselection if preselection else FilteredElementCollector(doc, curview.Id) \
    .WhereElementIsNotElementType() \
    .ToElements()

if not fab_elements:
    MessageBox.Show('No Fabrication parts found.')
    import sys
    sys.exit()

# Build property options with all 18 properties (excluding 'Name' and 'Comments' for special handling)
property_options = {}
for prop in ['CID', 'ServiceType', 'Service Name', 'Service Abbreviation', 'Size',
             'STRATUS Assembly', 'Line Number', 'STRATUS Status', 'Reference Level',
             'Item Number', 'Bundle Number', 'REF BS Designation', 'REF Line Number',
             'Specification', 'Hanger Rod Size', 'Valve Number']:
    try:
        values = set(filter(None, [get_property_value(elem, prop) for elem in fab_elements]))
        if values:
            property_options[prop] = sorted(values)
    except Exception, e:
        pass

# Handle 'Name' and 'Comments' separately with all elements
try:
    name_values = set(filter(None, [get_property_value(elem, 'Name') for elem in all_elements]))
    if name_values:
        property_options['Name'] = sorted(name_values)
except Exception, e:
    pass

try:
    comments_values = set(filter(None, [get_property_value(elem, 'Comments') for elem in all_elements]))
    if comments_values:
        property_options['Comments'] = sorted(comments_values)
except Exception, e:
    pass

if not property_options:
    MessageBox.Show('No properties found for the selected elements.')
    import sys
    sys.exit()

# Show form and process results
form = MultiPropertyFilterForm(property_options)
if form.ShowDialog() == DialogResult.OK and form.selected_filters:
    filtered_ids = []
    # Choose element set based on whether 'Name' or 'Comments' is in the filters
    elements_to_filter = all_elements if ('Name' in form.selected_filters or 'Comments' in form.selected_filters) else fab_elements
    for elem in elements_to_filter:
        matches = []
        for prop, filter_list in form.selected_filters.items():
            elem_value = get_property_value(elem, prop)
            # Check if elem_value matches any filter condition for this property
            prop_matches = []
            for values, is_and in filter_list:
                prop_matches.append((elem_value in values, is_and))
            # Combine matches for this property: AND requires all true, OR requires any true
            and_matches_prop = [m for m, is_and in prop_matches if is_and]
            or_matches_prop = [m for m, is_and in prop_matches if not is_and]
            prop_result = (not and_matches_prop or all(and_matches_prop)) and \
                         (not or_matches_prop or any(or_matches_prop))
            matches.append(prop_result)
        
        # Element must match all property conditions
        if all(matches):
            filtered_ids.append(elem.Id)

    revit.get_selection().set_to(filtered_ids)