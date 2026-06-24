# -*- coding: utf-8 -*-
import os
import clr
import System

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation
from System.Windows.Controls import (
    Grid,
    StackPanel,
    TextBlock,
    ComboBox,
    Button,
    RowDefinition,
    ColumnDefinition,
    Orientation
)
from System.Windows.Interop import WindowInteropHelper

from Autodesk.Revit.DB import (
    Transaction,
    FilteredElementCollector,
    StorageType,
    ElementId,
    ElementCategoryFilter
)
from Autodesk.Revit.UI import UIApplication, TaskDialog


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
uiapp = UIApplication(doc.Application)

FOLDER_NAME = r"C:\Temp"
FILE_PATH = os.path.join(FOLDER_NAME, "Ribbon_ParamToParam.txt")


# -----------------------------------------------------------------------------
# Settings file helpers
# -----------------------------------------------------------------------------
def ensure_settings_file():
    if not os.path.exists(FOLDER_NAME):
        os.makedirs(FOLDER_NAME)

    if not os.path.exists(FILE_PATH):
        with open(FILE_PATH, "w") as f:
            f.write("category=\ndatatype=\nsource=\ntarget=\n")


def read_previous_settings():
    ensure_settings_file()

    data = {
        "category": "",
        "datatype": "",
        "source": "",
        "target": ""
    }

    try:
        with open(FILE_PATH, "r") as f:
            for line in f:
                line = line.strip()
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip().lower()
                value = value.strip()
                if key in data:
                    data[key] = value
    except Exception:
        pass

    return data


def write_previous_settings(category, datatype, source, target):
    ensure_settings_file()

    try:
        with open(FILE_PATH, "w") as f:
            f.write("category={}\n".format(category or ""))
            f.write("datatype={}\n".format(datatype or ""))
            f.write("source={}\n".format(source or ""))
            f.write("target={}\n".format(target or ""))
    except Exception as e:
        TaskDialog.Show("Warning", "Could not save settings:\n{}".format(str(e)))


# -----------------------------------------------------------------------------
# Revit helpers
# -----------------------------------------------------------------------------
def get_revit_window_handle():
    try:
        return uidoc.Application.MainWindowHandle
    except Exception:
        try:
            return uiapp.MainWindowHandle
        except Exception:
            return System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle


def get_category_mapping():
    categories = {}

    try:
        for cat in doc.Settings.Categories:
            try:
                if cat and cat.Name and cat.Id and cat.Id != ElementId.InvalidElementId:
                    categories[cat.Name] = cat.Id
            except Exception:
                pass
    except Exception:
        pass

    return categories


def get_elements_by_category(category_name, category_map):
    cat_id = category_map.get(category_name)
    if cat_id is None:
        return []

    try:
        return list(
            FilteredElementCollector(doc)
            .WherePasses(ElementCategoryFilter(cat_id))
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []


def storage_type_to_label(storage_type):
    if storage_type == StorageType.String:
        return "String"
    elif storage_type == StorageType.Integer:
        return "Integer"
    elif storage_type == StorageType.Double:
        return "Double"
    elif storage_type == StorageType.ElementId:
        return "ElementId"
    return ""


def label_to_storage_type(label):
    if label == "String":
        return StorageType.String
    elif label == "Integer":
        return StorageType.Integer
    elif label == "Double":
        return StorageType.Double
    elif label == "ElementId":
        return StorageType.ElementId
    return None


def get_parameter_text(param):
    if not param:
        return ""

    try:
        if param.StorageType == StorageType.String:
            return param.AsString() or ""

        value = param.AsValueString()
        if value not in [None, ""]:
            return value

        if param.StorageType == StorageType.Integer:
            return str(param.AsInteger())

        if param.StorageType == StorageType.Double:
            return str(param.AsDouble())

        if param.StorageType == StorageType.ElementId:
            eid = param.AsElementId()
            if eid and eid != ElementId.InvalidElementId:
                return str(eid.IntegerValue)

    except Exception:
        pass

    return ""


def parameter_has_data(param):
    if not param:
        return False

    try:
        st = param.StorageType

        if st == StorageType.String:
            return (param.AsString() or "") != ""

        elif st == StorageType.Integer:
            return True

        elif st == StorageType.Double:
            return True

        elif st == StorageType.ElementId:
            eid = param.AsElementId()
            return eid and eid != ElementId.InvalidElementId

    except Exception:
        pass

    return False


def get_available_data_types(elements, sample_size=20):
    source_types = set()
    target_types = set()
    sample = elements[:min(sample_size, len(elements))]

    for elem in sample:
        for param in elem.Parameters:
            try:
                if not param or not param.Definition:
                    continue

                st = param.StorageType

                if st in [StorageType.String, StorageType.Integer, StorageType.Double, StorageType.ElementId]:
                    if parameter_has_data(param):
                        source_types.add(st)

                    if not param.IsReadOnly:
                        target_types.add(st)
            except Exception:
                pass

    matched = source_types.intersection(target_types)

    ordered = []
    for st in [StorageType.String, StorageType.Integer, StorageType.Double, StorageType.ElementId]:
        if st in matched:
            ordered.append(storage_type_to_label(st))

    return ordered


def get_source_parameters_with_data_by_type(elements, desired_type, sample_size=20):
    names = set()
    sample = elements[:min(sample_size, len(elements))]

    for elem in sample:
        for param in elem.Parameters:
            try:
                if (
                    param
                    and param.Definition
                    and param.StorageType == desired_type
                    and parameter_has_data(param)
                ):
                    names.add(param.Definition.Name)
            except Exception:
                pass

    return sorted(names)


def get_target_parameters_by_type(elements, desired_type, sample_size=20):
    names = set()
    sample = elements[:min(sample_size, len(elements))]

    for elem in sample:
        for param in elem.Parameters:
            try:
                if (
                    param
                    and param.Definition
                    and param.StorageType == desired_type
                    and not param.IsReadOnly
                ):
                    names.add(param.Definition.Name)
            except Exception:
                pass

    return sorted(names)


# -----------------------------------------------------------------------------
# Window
# -----------------------------------------------------------------------------
class Params2ParamWindow(Window):
    def __init__(self):
        Window.__init__(self)

        self.category_map = get_category_mapping()
        self.elements = []
        self.previous = read_previous_settings()

        self.Title = "Param -> Param"
        self.Width = 700
        self.Height = 285
        self.Topmost = True
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        self.build_ui()
        self.load_categories()
        self.apply_previous_settings()

        try:
            WindowInteropHelper(self).Owner = get_revit_window_handle()
        except Exception:
            pass

    def build_ui(self):
        root = Grid()
        root.Margin = Thickness(15)

        root.RowDefinitions.Add(RowDefinition())  # category
        root.RowDefinitions.Add(RowDefinition())  # datatype
        root.RowDefinitions.Add(RowDefinition())  # source/target
        root.RowDefinitions.Add(RowDefinition())  # buttons

        # Category
        category_panel = StackPanel()
        category_panel.Margin = Thickness(0, 0, 0, 10)

        category_label = TextBlock()
        category_label.Text = "Category"
        category_label.Margin = Thickness(0, 0, 0, 5)

        self.category_combo = ComboBox()
        self.category_combo.MinWidth = 250
        self.category_combo.SelectionChanged += self.category_changed

        category_panel.Children.Add(category_label)
        category_panel.Children.Add(self.category_combo)

        Grid.SetRow(category_panel, 0)
        root.Children.Add(category_panel)

        # Data type
        datatype_panel = StackPanel()
        datatype_panel.Margin = Thickness(0, 0, 0, 10)

        datatype_label = TextBlock()
        datatype_label.Text = "Data Type"
        datatype_label.Margin = Thickness(0, 0, 0, 5)

        self.datatype_combo = ComboBox()
        self.datatype_combo.MinWidth = 250
        self.datatype_combo.SelectionChanged += self.datatype_changed

        datatype_panel.Children.Add(datatype_label)
        datatype_panel.Children.Add(self.datatype_combo)

        Grid.SetRow(datatype_panel, 1)
        root.Children.Add(datatype_panel)

        # Mapping row
        mapping_grid = Grid()
        mapping_grid.Margin = Thickness(0, 0, 0, 10)
        mapping_grid.ColumnDefinitions.Add(ColumnDefinition())
        mapping_grid.ColumnDefinitions.Add(ColumnDefinition())

        # Source
        source_panel = StackPanel()
        source_panel.Margin = Thickness(0, 0, 10, 0)

        source_label = TextBlock()
        source_label.Text = "Parameter with data"
        source_label.Margin = Thickness(0, 0, 0, 5)

        self.source_combo = ComboBox()
        self.source_combo.MinWidth = 250

        source_panel.Children.Add(source_label)
        source_panel.Children.Add(self.source_combo)

        Grid.SetColumn(source_panel, 0)
        mapping_grid.Children.Add(source_panel)

        # Target
        target_panel = StackPanel()
        target_panel.Margin = Thickness(10, 0, 0, 0)

        target_label = TextBlock()
        target_label.Text = "Push data to"
        target_label.Margin = Thickness(0, 0, 0, 5)

        self.target_combo = ComboBox()
        self.target_combo.MinWidth = 250

        target_panel.Children.Add(target_label)
        target_panel.Children.Add(self.target_combo)

        Grid.SetColumn(target_panel, 1)
        mapping_grid.Children.Add(target_panel)

        Grid.SetRow(mapping_grid, 2)
        root.Children.Add(mapping_grid)

        # Buttons
        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Right

        run_button = Button()
        run_button.Content = "Run"
        run_button.Width = 90
        run_button.Height = 25
        run_button.Margin = Thickness(0, 0, 10, 0)
        run_button.Click += self.run_click

        cancel_button = Button()
        cancel_button.Content = "Cancel"
        cancel_button.Width = 90
        cancel_button.Height = 25
        cancel_button.Click += self.cancel_click

        button_panel.Children.Add(run_button)
        button_panel.Children.Add(cancel_button)

        Grid.SetRow(button_panel, 3)
        root.Children.Add(button_panel)

        self.Content = root

    def load_categories(self):
        self.category_combo.Items.Clear()
        for name in sorted(self.category_map.keys()):
            self.category_combo.Items.Add(name)

    def apply_previous_settings(self):
        saved_category = self.previous.get("category", "")
        saved_datatype = self.previous.get("datatype", "")
        saved_source = self.previous.get("source", "")
        saved_target = self.previous.get("target", "")

        if saved_category:
            for item in self.category_combo.Items:
                if str(item) == saved_category:
                    self.category_combo.SelectedItem = item
                    break

        if self.category_combo.SelectedItem:
            self.load_data_types_for_category(str(self.category_combo.SelectedItem))

            if saved_datatype:
                for item in self.datatype_combo.Items:
                    if str(item) == saved_datatype:
                        self.datatype_combo.SelectedItem = item
                        break

            if self.datatype_combo.SelectedItem:
                self.load_parameters_for_type(str(self.datatype_combo.SelectedItem))

                if saved_source:
                    for item in self.source_combo.Items:
                        if str(item) == saved_source:
                            self.source_combo.SelectedItem = item
                            break

                if saved_target:
                    for item in self.target_combo.Items:
                        if str(item) == saved_target:
                            self.target_combo.SelectedItem = item
                            break

    def load_data_types_for_category(self, category_name):
        self.datatype_combo.Items.Clear()
        self.source_combo.Items.Clear()
        self.target_combo.Items.Clear()
        self.elements = []

        if not category_name:
            return

        self.elements = get_elements_by_category(category_name, self.category_map)

        if not self.elements:
            return

        data_types = get_available_data_types(self.elements)

        for name in data_types:
            self.datatype_combo.Items.Add(name)

    def load_parameters_for_type(self, datatype_label):
        self.source_combo.Items.Clear()
        self.target_combo.Items.Clear()

        if not datatype_label or not self.elements:
            return

        storage_type = label_to_storage_type(datatype_label)
        if storage_type is None:
            return

        source_params = get_source_parameters_with_data_by_type(self.elements, storage_type)
        target_params = get_target_parameters_by_type(self.elements, storage_type)

        for name in source_params:
            self.source_combo.Items.Add(name)

        for name in target_params:
            self.target_combo.Items.Add(name)

    def category_changed(self, sender, args):
        category_name = self.category_combo.SelectedItem
        if not category_name:
            return

        self.load_data_types_for_category(str(category_name))

    def datatype_changed(self, sender, args):
        datatype_label = self.datatype_combo.SelectedItem
        if not datatype_label:
            return

        self.load_parameters_for_type(str(datatype_label))

    def run_click(self, sender, args):
        category_name = self.category_combo.SelectedItem
        datatype_label = self.datatype_combo.SelectedItem
        source_name = self.source_combo.SelectedItem
        target_name = self.target_combo.SelectedItem

        if not category_name:
            TaskDialog.Show("Warning", "Please select a category.")
            return

        if not self.elements:
            TaskDialog.Show("Warning", "No elements found in the selected category.")
            return

        if not datatype_label:
            TaskDialog.Show("Warning", "Please select a data type.")
            return

        if not source_name:
            TaskDialog.Show("Warning", "Please select a source parameter.")
            return

        if not target_name:
            TaskDialog.Show("Warning", "Please select a target parameter.")
            return

        category_name = str(category_name)
        datatype_label = str(datatype_label)
        source_name = str(source_name)
        target_name = str(target_name)

        selected_type = label_to_storage_type(datatype_label)
        if selected_type is None:
            TaskDialog.Show("Warning", "Invalid data type selected.")
            return

        updated = 0
        failed = 0
        t = None

        try:
            t = Transaction(doc, "Copy parameter values")
            t.Start()

            for elem in self.elements:
                try:
                    source_param = elem.LookupParameter(source_name)
                    target_param = elem.LookupParameter(target_name)

                    if not source_param or not target_param:
                        failed += 1
                        continue

                    if target_param.IsReadOnly:
                        failed += 1
                        continue

                    if source_param.StorageType != selected_type:
                        failed += 1
                        continue

                    if target_param.StorageType != selected_type:
                        failed += 1
                        continue

                    if selected_type == StorageType.String:
                        target_param.Set(source_param.AsString() or "")
                        updated += 1

                    elif selected_type == StorageType.Integer:
                        target_param.Set(source_param.AsInteger())
                        updated += 1

                    elif selected_type == StorageType.Double:
                        target_param.Set(source_param.AsDouble())
                        updated += 1

                    elif selected_type == StorageType.ElementId:
                        eid = source_param.AsElementId()
                        if eid and eid != ElementId.InvalidElementId:
                            target_param.Set(eid)
                            updated += 1
                        else:
                            failed += 1

                    else:
                        failed += 1

                except Exception:
                    failed += 1

            t.Commit()

        except Exception as e:
            try:
                if t:
                    t.RollBack()
            except Exception:
                pass

            TaskDialog.Show("Error", "Error:\n{}".format(str(e)))
            return

        write_previous_settings(category_name, datatype_label, source_name, target_name)

        msg = "{} elements updated.".format(updated)
        if failed:
            msg += "\n{} elements failed.".format(failed)

        TaskDialog.Show("Results", msg)
        self.Close()

    def cancel_click(self, sender, args):
        self.Close()


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
def main():
    window = Params2ParamWindow()
    window.ShowDialog()


if __name__ == "__main__":
    main()