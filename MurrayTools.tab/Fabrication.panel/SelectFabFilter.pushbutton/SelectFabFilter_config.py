# -*- coding: utf-8 -*-
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Form, Label, ComboBox, Button, ListBox, 
                                CheckBox, DialogResult, FormBorderStyle, FormStartPosition, 
                                SelectionMode, Control, ComboBoxStyle, TextBox)
from System import Array
from System.Drawing import Point, Size
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, FabricationConfiguration
from pyrevit import revit, forms
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import (get_parameter_value_by_name_AsString, 
                                     get_parameter_value_by_name_AsValueString, 
                                     get_parameter_value_by_name_AsInteger)

Shared_Params()

# Define the active Revit document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

class MultiPropertyFilterForm(Form):
    def __init__(self, property_options):
        self.property_options = property_options
        self.InitializeComponents()
        self.selected_filters = {}  # {property: (values, is_and)}
        if self.property_options:
            first_property = self.property_combo.Items[0]
            if first_property in self.property_options:
                self.property_combo.SelectedItem = first_property
            self.property_changed(None, None)
        self.update_filter_display()

    def InitializeComponents(self):
        self.Text = "Multi-Property Filter"
        self.Size = Size(500, 550)  # Your chosen size
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        # Property selection
        self.property_label = Label()
        self.property_label.Text = "Select Property:"
        self.property_label.Location = Point(10, 10)
        self.property_label.Size = Size(100, 20)

        self.property_combo = ComboBox()
        self.property_combo.Location = Point(120, 10)
        self.property_combo.Size = Size(150, 20)
        self.property_combo.DropDownStyle = ComboBoxStyle.DropDownList
        properties = list(self.property_options.keys())
        if properties:
            self.property_combo.Items.AddRange(Array[object](properties))
        self.property_combo.SelectedIndexChanged += self.property_changed
        self.Controls.Add(self.property_combo)

        # AND/OR toggle (right of dropdown)
        self.logic_check = CheckBox()
        self.logic_check.Text = "AND logic (unchecked = OR)"
        self.logic_check.Location = Point(275, 10)
        self.logic_check.Size = Size(200, 20)

        # Search label
        self.search_label = Label()
        self.search_label.Text = "Search:"
        self.search_label.Location = Point(10, 40)
        self.search_label.Size = Size(100, 20)

        # Search bar
        self.search_box = TextBox()
        self.search_box.Location = Point(120, 40)
        self.search_box.Size = Size(300, 20)
        self.search_box.TextChanged += self.search_changed
        self.Controls.Add(self.search_box)

        # Values list
        self.values_label = Label()
        self.values_label.Text = "Select Values:"
        self.values_label.Location = Point(10, 70)
        self.values_label.Size = Size(100, 20)

        # Add Filter button (under "Select Values" label)
        self.add_button = Button()
        self.add_button.Text = "Add Filter"
        self.add_button.Location = Point(12, 90)  # Below label (70 + 20)
        self.add_button.Size = Size(80, 23)
        self.add_button.Click += self.add_filter

        self.values_list = ListBox()
        self.values_list.Location = Point(120, 70)
        self.values_list.Size = Size(300, 270)
        self.values_list.SelectionMode = SelectionMode.MultiExtended

        # Filter feedback label (under ListBox)
        self.filter_display = Label()
        self.filter_display.Text = "No filters added yet."
        self.filter_display.Location = Point(10, 350)
        self.filter_display.Size = Size(300, 50)
        self.filter_display.AutoSize = False

        # Buttons (centered at bottom)
        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.Location = Point(165, 480)
        self.ok_button.Size = Size(80, 23)
        self.ok_button.Click += self.ok_clicked

        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Location = Point(255, 480)  # 165 + 80 + 10
        self.cancel_button.Size = Size(80, 23)
        self.cancel_button.Click += self.cancel_clicked

        self.Controls.AddRange(Array[Control]([
            self.property_label, 
            self.search_label,  # Added search label
            self.values_label, 
            self.values_list,
            self.logic_check, 
            self.add_button,
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
            self.selected_filters[selected_property] = (selected_values, self.logic_check.Checked)
            self.update_filter_display()

    def update_filter_display(self, sender=None, args=None):
        if not self.selected_filters:
            self.filter_display.Text = "No filters added yet."
        else:
            filter_text = "Filters:\n"
            for prop, (values, is_and) in self.selected_filters.items():
                mode = "AND" if is_and else "OR"
                filter_text += "%s (%s): %s\n" % (prop, mode, ", ".join(str(v) for v in values))
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
elements = preselection if preselection else FilteredElementCollector(doc, curview.Id) \
    .OfClass(DB.FabricationPart) \
    .WhereElementIsNotElementType() \
    .ToElements()

if not elements:
    forms.alert('No Fabrication parts found.', exitscript=True)

# Build property options with all 18 properties
property_options = {}
for prop in ['CID', 'ServiceType', 'Name', 'Service Name', 'Service Abbreviation', 'Size',
             'STRATUS Assembly', 'Line Number', 'STRATUS Status', 'Reference Level',
             'Item Number', 'Bundle Number', 'REF BS Designation', 'REF Line Number',
             'Comments', 'Specification', 'Hanger Rod Size', 'Valve Number']:
    try:
        values = set(filter(None, [get_property_value(elem, prop) for elem in elements]))
        if values:
            property_options[prop] = sorted(values)
    except Exception, e:
        pass

if not property_options:
    forms.alert('No properties found for the selected elements.', exitscript=True)

# Show form and process results
form = MultiPropertyFilterForm(property_options)
if form.ShowDialog() == DialogResult.OK and form.selected_filters:
    filtered_ids = []
    for elem in elements:
        matches = []
        for prop, (values, is_and) in form.selected_filters.items():
            elem_value = get_property_value(elem, prop)
            matches.append((elem_value in values, is_and))
        
        and_matches = [m for m, is_and in matches if is_and]
        or_matches = [m for m, is_and in matches if not is_and]
        if (not and_matches or all(and_matches)) and (not or_matches or any(or_matches)):
            filtered_ids.append(elem.Id)

    revit.get_selection().set_to(filtered_ids)