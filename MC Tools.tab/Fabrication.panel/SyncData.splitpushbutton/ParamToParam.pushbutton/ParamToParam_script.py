# -*- coding: utf-8 -*-
import os
import clr
import json
import System

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Windows import (
    Window,
    Thickness,
    HorizontalAlignment,
    WindowStartupLocation,
    ResizeMode,
    TextWrapping,
    FontWeights
)
from System.Windows.Controls import (
    Grid,
    StackPanel,
    TextBlock,
    ComboBox,
    Button,
    RowDefinition,
    ColumnDefinition,
    Orientation,
    ScrollViewer,
    Border,
    DockPanel,
    Dock
)
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Media import Brushes

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
# TXT settings helpers
# -----------------------------------------------------------------------------
def ensure_settings_file():
    if not os.path.exists(FOLDER_NAME):
        os.makedirs(FOLDER_NAME)

    if not os.path.exists(FILE_PATH):
        with open(FILE_PATH, "w") as f:
            f.write("{}")


def load_saved_settings():
    ensure_settings_file()

    try:
        with open(FILE_PATH, "r") as f:
            text = f.read().strip()

        if not text:
            return {}

        data = json.loads(text)
        return normalize_loaded_settings(data)

    except Exception:
        return {}


def save_settings(settings):
    ensure_settings_file()

    try:
        with open(FILE_PATH, "w") as f:
            f.write(json.dumps(settings, indent=2, sort_keys=True))
        return True, None
    except Exception as e:
        return False, str(e)


def normalize_loaded_settings(data):
    settings = {}

    if not isinstance(data, dict):
        return settings

    for category_name, rules in data.items():
        try:
            cat = str(category_name or "").strip()
        except Exception:
            cat = ""

        if not cat:
            continue

        cleaned_rules = []

        if isinstance(rules, list):
            for rule in rules:
                if not isinstance(rule, dict):
                    continue

                cleaned_rules.append({
                    "datatype": str(rule.get("datatype", "") or "").strip(),
                    "source": str(rule.get("source", "") or "").strip(),
                    "target": str(rule.get("target", "") or "").strip()
                })

        settings[cat] = cleaned_rules

    return settings


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
                if st not in [StorageType.String, StorageType.Integer, StorageType.Double, StorageType.ElementId]:
                    continue

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
class ParamsToParamConfigWindow(Window):
    def __init__(self):
        Window.__init__(self)

        self.category_map = get_category_mapping()
        self.settings = load_saved_settings()
        self.clipboard_rules = None

        self.Title = "Param -> Param Configuration Manager"
        self.Width = 1050
        self.Height = 750
        self.Topmost = True
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        self.build_ui()
        self.refresh_buckets()

        try:
            WindowInteropHelper(self).Owner = get_revit_window_handle()
        except Exception:
            pass

    def build_ui(self):
        root = Grid()
        root.Margin = Thickness(15)

        root.RowDefinitions.Add(RowDefinition())  # info & top add bar
        root.RowDefinitions.Add(RowDefinition())  # categories scroll viewer
        root.RowDefinitions.Add(RowDefinition())  # bottom global buttons

        root.RowDefinitions[0].Height = System.Windows.GridLength(75)
        root.RowDefinitions[1].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        root.RowDefinitions[2].Height = System.Windows.GridLength(45)

        # Top Bar: Info text + Category Add selector
        top_panel = StackPanel()
        
        info = TextBlock()
        info.Text = (
            "Configure parameter-to-parameter copy rules grouped by category buckets. "
            "Settings save automatically to C:\\Temp\\Ribbon_ParamToParam.txt."
        )
        info.TextWrapping = TextWrapping.Wrap
        info.Margin = Thickness(0, 0, 0, 8)
        top_panel.Children.Add(info)

        add_bar_grid = Grid()
        add_bar_grid.ColumnDefinitions.Add(ColumnDefinition())
        add_bar_grid.ColumnDefinitions.Add(ColumnDefinition())
        add_bar_grid.ColumnDefinitions[0].Width = System.Windows.GridLength(350)
        add_bar_grid.ColumnDefinitions[1].Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)

        add_controls_stack = StackPanel()
        add_controls_stack.Orientation = Orientation.Horizontal

        self.all_cats_combo = ComboBox()
        self.all_cats_combo.Width = 240
        self.all_cats_combo.Height = 26
        self.all_cats_combo.Margin = Thickness(0, 0, 8, 0)
        add_controls_stack.Children.Add(self.all_cats_combo)

        self.add_cat_button = Button()
        self.add_cat_button.Content = "+ Add Category"
        self.add_cat_button.Width = 100
        self.add_cat_button.Height = 26
        self.add_cat_button.Click += self.add_category_click
        add_controls_stack.Children.Add(self.add_cat_button)

        Grid.SetColumn(add_controls_stack, 0)
        add_bar_grid.Children.Add(add_controls_stack)

        top_panel.Children.Add(add_bar_grid)
        Grid.SetRow(top_panel, 0)
        root.Children.Add(top_panel)

        # Populate available categories dropdown
        self.all_cats_combo.Items.Clear()
        for cat_name in sorted(self.category_map.keys()):
            self.all_cats_combo.Items.Add(cat_name)
        if self.all_cats_combo.Items.Count > 0:
            self.all_cats_combo.SelectedIndex = 0

        # Middle: Scrollable Buckets Container
        self.buckets_scroll = ScrollViewer()
        self.buckets_scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        self.buckets_scroll.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Disabled
        self.buckets_scroll.Margin = Thickness(0, 5, 0, 5)

        self.buckets_panel = StackPanel()
        self.buckets_scroll.Content = self.buckets_panel

        Grid.SetRow(self.buckets_scroll, 1)
        root.Children.Add(self.buckets_scroll)

        # Bottom Global Action Panel
        bottom_panel = StackPanel()
        bottom_panel.Orientation = Orientation.Horizontal
        bottom_panel.HorizontalAlignment = HorizontalAlignment.Right
        bottom_panel.Margin = Thickness(0, 8, 0, 0)

        self.run_all_button = Button()
        self.run_all_button.Content = "Run All Categories"
        self.run_all_button.Width = 135
        self.run_all_button.Height = 28
        self.run_all_button.Margin = Thickness(0, 0, 10, 0)
        self.run_all_button.Click += self.run_all_click

        self.close_button = Button()
        self.close_button.Content = "Close"
        self.close_button.Width = 90
        self.close_button.Height = 28
        self.close_button.Click += self.close_click

        bottom_panel.Children.Add(self.run_all_button)
        bottom_panel.Children.Add(self.close_button)

        Grid.SetRow(bottom_panel, 2)
        root.Children.Add(bottom_panel)

        self.Content = root

    def add_category_click(self, sender, args):
        selected = self.all_cats_combo.SelectedItem
        if not selected:
            return
        cat_name = str(selected)
        if cat_name not in self.settings:
            self.settings[cat_name] = [{"datatype": "", "source": "", "target": ""}]
            save_settings(self.settings)
            self.refresh_buckets()

    def refresh_buckets(self):
        self.buckets_panel.Children.Clear()

        if not self.settings:
            empty_block = TextBlock()
            empty_block.Text = "No categories configured yet. Select a category above and click '+ Add Category'."
            empty_block.FontStyle = System.Windows.FontStyles.Italic
            empty_block.Margin = Thickness(5, 10, 5, 5)
            self.buckets_panel.Children.Add(empty_block)
            return

        for cat_name in sorted(self.settings.keys()):
            rules = self.settings.get(cat_name, [])
            bucket_border = self.create_category_bucket(cat_name, rules)
            self.buckets_panel.Children.Add(bucket_border)

    def create_category_bucket(self, cat_name, rules):
        elements = get_elements_by_category(cat_name, self.category_map)
        data_types = get_available_data_types(elements)

        card_border = Border()
        card_border.BorderBrush = Brushes.LightGray
        card_border.BorderThickness = Thickness(1)
        card_border.Margin = Thickness(0, 0, 0, 12)
        card_border.Padding = Thickness(10)
        card_border.Background = Brushes.White

        card_stack = StackPanel()

        # Bucket Header
        header_panel = DockPanel()
        header_panel.Margin = Thickness(0, 0, 0, 6)

        title_block = TextBlock()
        title_block.Text = "{}  (Elements: {})".format(cat_name.upper(), len(elements))
        title_block.FontWeight = FontWeights.Bold
        title_block.FontSize = 13
        DockPanel.SetDock(title_block, Dock.Left)
        header_panel.Children.Add(title_block)

        btns_stack = StackPanel()
        btns_stack.Orientation = Orientation.Horizontal
        btns_stack.HorizontalAlignment = HorizontalAlignment.Right

        copy_btn = Button()
        copy_btn.Content = "Copy"
        copy_btn.Width = 50
        copy_btn.Height = 24
        copy_btn.Margin = Thickness(0, 0, 5, 0)
        copy_btn.Click += lambda s, e, cn=cat_name: self.copy_category_rules(cn)

        paste_btn = Button()
        paste_btn.Content = "Paste"
        paste_btn.Width = 50
        paste_btn.Height = 24
        paste_btn.Margin = Thickness(0, 0, 5, 0)
        paste_btn.Click += lambda s, e, cn=cat_name: self.paste_category_rules(cn)

        add_rule_btn = Button()
        add_rule_btn.Content = "+ Add Rule"
        add_rule_btn.Width = 75
        add_rule_btn.Height = 24
        add_rule_btn.Margin = Thickness(0, 0, 5, 0)
        add_rule_btn.Click += lambda s, e, cn=cat_name, cs=card_stack: self.add_rule_to_bucket(cn, cs)

        remove_cat_btn = Button()
        remove_cat_btn.Content = "X"
        remove_cat_btn.Width = 24
        remove_cat_btn.Height = 24
        remove_cat_btn.Foreground = Brushes.Red
        remove_cat_btn.FontWeight = FontWeights.Bold
        remove_cat_btn.Click += lambda s, e, cn=cat_name: self.remove_category_config(cn)

        btns_stack.Children.Add(copy_btn)
        btns_stack.Children.Add(paste_btn)
        btns_stack.Children.Add(add_rule_btn)
        btns_stack.Children.Add(remove_cat_btn)

        DockPanel.SetDock(btns_stack, Dock.Right)
        header_panel.Children.Add(btns_stack)
        card_stack.Children.Add(header_panel)

        # Table Column Headers for rules inside bucket
        header_grid = Grid()
        header_grid.Margin = Thickness(0, 0, 0, 3)
        header_grid.ColumnDefinitions.Add(ColumnDefinition())
        header_grid.ColumnDefinitions.Add(ColumnDefinition())
        header_grid.ColumnDefinitions.Add(ColumnDefinition())
        col_del = ColumnDefinition()
        col_del.Width = System.Windows.GridLength(35)
        header_grid.ColumnDefinitions.Add(col_del)

        def make_col_header(text, col):
            tb = TextBlock()
            tb.Text = text
            tb.FontWeight = FontWeights.Bold
            tb.FontSize = 11
            tb.Margin = Thickness(3, 0, 3, 0)
            Grid.SetColumn(tb, col)
            return tb

        header_grid.Children.Add(make_col_header("Data Type", 0))
        header_grid.Children.Add(make_col_header("Source Parameter", 1))
        header_grid.Children.Add(make_col_header("Target Parameter", 2))
        header_grid.Children.Add(make_col_header("", 3))
        card_stack.Children.Add(header_grid)

        # Rules Rows Container inside bucket
        rules_stack = StackPanel()
        card_stack.Children.Add(rules_stack)

        if not rules:
            rules = [{"datatype": "", "source": "", "target": ""}]

        for rule in rules:
            self.build_rule_row(rules_stack, cat_name, elements, data_types, rule)

        card_border.Child = card_stack
        return card_border

    def build_rule_row(self, parent_panel, cat_name, elements, data_types, rule_data):
        row_grid = Grid()
        row_grid.Margin = Thickness(0, 0, 0, 4)
        row_grid.ColumnDefinitions.Add(ColumnDefinition())
        row_grid.ColumnDefinitions.Add(ColumnDefinition())
        row_grid.ColumnDefinitions.Add(ColumnDefinition())
        col_del = ColumnDefinition()
        col_del.Width = System.Windows.GridLength(35)
        row_grid.ColumnDefinitions.Add(col_del)

        datatype_combo = ComboBox()
        datatype_combo.Margin = Thickness(2, 0, 2, 0)
        datatype_combo.Height = 24
        Grid.SetColumn(datatype_combo, 0)
        row_grid.Children.Add(datatype_combo)

        source_combo = ComboBox()
        source_combo.Margin = Thickness(2, 0, 2, 0)
        source_combo.Height = 24
        Grid.SetColumn(source_combo, 1)
        row_grid.Children.Add(source_combo)

        target_combo = ComboBox()
        target_combo.Margin = Thickness(2, 0, 2, 0)
        target_combo.Height = 24
        Grid.SetColumn(target_combo, 2)
        row_grid.Children.Add(target_combo)

        delete_btn = Button()
        delete_btn.Content = "X"
        delete_btn.Width = 24
        delete_btn.Height = 22
        delete_btn.Foreground = Brushes.Red
        delete_btn.FontWeight = FontWeights.Bold
        Grid.SetColumn(delete_btn, 3)
        row_grid.Children.Add(delete_btn)

        # Populate data types
        self.fill_combo_items(datatype_combo, data_types, rule_data.get("datatype", ""))

        def update_param_combos():
            dt_label = str(datatype_combo.SelectedItem or "")
            st = label_to_storage_type(dt_label)

            source_combo.Items.Clear()
            target_combo.Items.Clear()
            source_combo.Items.Add("")
            target_combo.Items.Add("")
            source_combo.SelectedIndex = 0
            target_combo.SelectedIndex = 0

            if st is not None and elements:
                src_params = get_source_parameters_with_data_by_type(elements, st)
                tgt_params = get_target_parameters_by_type(elements, st)
                for sp in src_params:
                    source_combo.Items.Add(sp)
                for tp in tgt_params:
                    target_combo.Items.Add(tp)

            if rule_data.get("source"):
                for item in source_combo.Items:
                    if str(item) == str(rule_data.get("source")):
                        source_combo.SelectedItem = item
                        break

            if rule_data.get("target"):
                for item in target_combo.Items:
                    if str(item) == str(rule_data.get("target")):
                        target_combo.SelectedItem = item
                        break

        update_param_combos()

        def datatype_changed(sender, args):
            rule_data["datatype"] = str(datatype_combo.SelectedItem or "")
            update_param_combos()
            self.save_bucket_rules(cat_name, parent_panel)

        datatype_combo.SelectionChanged += datatype_changed

        def param_changed(sender, args):
            rule_data["source"] = str(source_combo.SelectedItem or "")
            rule_data["target"] = str(target_combo.SelectedItem or "")
            self.save_bucket_rules(cat_name, parent_panel)

        source_combo.SelectionChanged += param_changed
        target_combo.SelectionChanged += param_changed

        def delete_row_click(sender, args):
            try:
                parent_panel.Children.Remove(row_grid)
            except Exception:
                pass
            self.save_bucket_rules(cat_name, parent_panel)

        delete_btn.Click += delete_row_click

        parent_panel.Children.Add(row_grid)

    def fill_combo_items(self, combo, items, selected_val):
        combo.Items.Clear()
        combo.Items.Add("")
        for item in items:
            combo.Items.Add(item)
        combo.SelectedIndex = 0
        if selected_val:
            for item in combo.Items:
                if str(item) == str(selected_val):
                    combo.SelectedItem = item
                    break

    def add_rule_to_bucket(self, cat_name, card_stack):
        if cat_name in self.settings:
            self.settings[cat_name].append({"datatype": "", "source": "", "target": ""})
            save_settings(self.settings)
            self.refresh_buckets()

    def save_bucket_rules(self, cat_name, rules_stack):
        rules = []
        for child in rules_stack.Children:
            if isinstance(child, Grid):
                dt, src, tgt = "", "", ""
                try:
                    if child.Children[0].SelectedItem:
                        dt = str(child.Children[0].SelectedItem).strip()
                    if child.Children[1].SelectedItem:
                        src = str(child.Children[1].SelectedItem).strip()
                    if child.Children[2].SelectedItem:
                        tgt = str(child.Children[2].SelectedItem).strip()
                except Exception:
                    pass
                rules.append({"datatype": dt, "source": src, "target": tgt})

        if rules:
            self.settings[cat_name] = rules
        else:
            if cat_name in self.settings:
                del self.settings[cat_name]
        save_settings(self.settings)

    def copy_category_rules(self, cat_name):
        rules = self.settings.get(cat_name, [])
        # Deep copy rules
        self.clipboard_rules = json.loads(json.dumps(rules))
        TaskDialog.Show("Copied", "Rules for '{}' copied to clipboard buffer.".format(cat_name))

    def paste_category_rules(self, cat_name):
        if not self.clipboard_rules:
            TaskDialog.Show("Warning", "No rules in clipboard buffer to paste. Copy a category first.")
            return
        self.settings[cat_name] = json.loads(json.dumps(self.clipboard_rules))
        save_settings(self.settings)
        self.refresh_buckets()
        TaskDialog.Show("Pasted", "Rules pasted to '{}'.".format(cat_name))

    def remove_category_config(self, cat_name):
        if cat_name in self.settings:
            del self.settings[cat_name]
            save_settings(self.settings)
            self.refresh_buckets()

    def run_all_click(self, sender, args):
        ok, err = save_settings(self.settings)
        if not ok:
            TaskDialog.Show("Error", "Could not save configurations before run:\n{}".format(err or "Unknown error"))
            return

        categories_to_run = []
        for cat_name, rules in sorted(self.settings.items()):
            if isinstance(rules, list):
                categories_to_run.append(cat_name)

        if not categories_to_run:
            TaskDialog.Show("Warning", "No saved category configurations found.")
            return

        self.run_categories(categories_to_run)

    def run_categories(self, categories_to_run):
        total_updated = 0
        total_failed = 0
        total_invalid_rules = 0
        categories_run = 0
        categories_skipped = 0
        summary_lines = []
        t = None

        try:
            t = Transaction(doc, "Copy parameter values (multi-rule)")
            t.Start()

            for category_name in categories_to_run:
                raw_rules = self.settings.get(category_name, [])
                executable_rules = []
                invalid_for_category = 0

                for raw_rule in raw_rules:
                    datatype_label = str(raw_rule.get("datatype", "") or "").strip()
                    source_name = str(raw_rule.get("source", "") or "").strip()
                    target_name = str(raw_rule.get("target", "") or "").strip()

                    if not datatype_label and not source_name and not target_name:
                        continue

                    if not datatype_label or not source_name or not target_name:
                        invalid_for_category += 1
                        total_invalid_rules += 1
                        continue

                    if source_name == target_name:
                        invalid_for_category += 1
                        total_invalid_rules += 1
                        continue

                    st = label_to_storage_type(datatype_label)
                    if st is None:
                        invalid_for_category += 1
                        total_invalid_rules += 1
                        continue

                    executable_rules.append({
                        "datatype": datatype_label,
                        "source": source_name,
                        "target": target_name,
                        "storage_type": st
                    })

                if not executable_rules:
                    categories_skipped += 1
                    summary_lines.append("{}: skipped (no valid rules)".format(category_name))
                    continue

                elements = get_elements_by_category(category_name, self.category_map)
                if not elements:
                    categories_skipped += 1
                    summary_lines.append("{}: skipped (no elements found)".format(category_name))
                    continue

                cat_updated = 0
                cat_failed = 0
                categories_run += 1

                for elem in elements:
                    for rule in executable_rules:
                        try:
                            source_param = elem.LookupParameter(rule["source"])
                            target_param = elem.LookupParameter(rule["target"])

                            if not source_param or not target_param:
                                cat_failed += 1
                                continue

                            if target_param.IsReadOnly:
                                cat_failed += 1
                                continue

                            if source_param.StorageType != rule["storage_type"]:
                                cat_failed += 1
                                continue

                            if target_param.StorageType != rule["storage_type"]:
                                cat_failed += 1
                                continue

                            if rule["storage_type"] == StorageType.String:
                                target_param.Set(source_param.AsString() or "")
                                cat_updated += 1

                            elif rule["storage_type"] == StorageType.Integer:
                                target_param.Set(source_param.AsInteger())
                                cat_updated += 1

                            elif rule["storage_type"] == StorageType.Double:
                                target_param.Set(source_param.AsDouble())
                                cat_updated += 1

                            elif rule["storage_type"] == StorageType.ElementId:
                                eid = source_param.AsElementId()
                                if eid and eid != ElementId.InvalidElementId:
                                    target_param.Set(eid)
                                    cat_updated += 1
                                else:
                                    cat_failed += 1

                            else:
                                cat_failed += 1

                        except Exception:
                            cat_failed += 1

                total_updated += cat_updated
                total_failed += cat_failed

                line = "{}: {} updated".format(category_name, cat_updated)
                if cat_failed:
                    line += ", {} failed".format(cat_failed)
                if invalid_for_category:
                    line += ", {} invalid rules skipped".format(invalid_for_category)

                summary_lines.append(line)

            t.Commit()

        except Exception as e:
            try:
                if t:
                    t.RollBack()
            except Exception:
                pass

            TaskDialog.Show("Error", "Run failed:\n{}".format(str(e)))
            return

        message = (
            "Categories run: {0}\n"
            "Categories skipped: {1}\n"
            "Total updated: {2}\n"
            "Total failed: {3}\n"
            "Invalid rules skipped: {4}\n\n"
            "{5}"
        ).format(
            categories_run,
            categories_skipped,
            total_updated,
            total_failed,
            total_invalid_rules,
            "\n".join(summary_lines) if summary_lines else "No category results."
        )

        TaskDialog.Show("Results", message)

    def close_click(self, sender, args):
        self.Close()


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
def main():
    if not uidoc or not doc:
        TaskDialog.Show("Error", "No active Revit document found.")
        return

    window = ParamsToParamConfigWindow()
    window.ShowDialog()


if __name__ == "__main__":
    main()