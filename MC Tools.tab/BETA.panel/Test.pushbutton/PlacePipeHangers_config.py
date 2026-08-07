# -*- coding: UTF-8 -*-
import os
import clr
import traceback
import System
import re

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System import Array
from System.Windows import (
    Window, Thickness, WindowStyle, ResizeMode,
    WindowStartupLocation, HorizontalAlignment, VerticalAlignment,
    GridLength, GridUnitType, TextAlignment, FontWeights, CornerRadius,
    TextWrapping
)
from System.Windows.Controls import (
    TextBox, Button, Grid, RowDefinition, ColumnDefinition, ComboBox,
    ScrollViewer, StackPanel, ScrollBarVisibility, TextBlock, Orientation, CheckBox, Border
)
from System.Windows.Media import FontFamily, SolidColorBrush, Colors, Brushes

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FabricationConfiguration
from Autodesk.Revit.UI import TaskDialog

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
RevitINT = float(app.VersionNumber)

# Per-project configuration path setup
file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)
if not file_name:
    file_name = doc.Title

FOLDER_NAME = r"c:\Temp"
project_name = file_name.replace(" ", "_")
FILEPATH = os.path.join(FOLDER_NAME, "Ribbon_Pipe-Hanger-Config_{}.txt".format(project_name))

if not os.path.exists(FOLDER_NAME):
    os.makedirs(FOLDER_NAME)


def get_hangers_for_service(doc, service_name):
    """Retrieves hangers available within a specific service, with '--- NONE ---' at the top."""
    names = set()
    try:
        config = FabricationConfiguration.GetFabricationConfiguration(doc)
        if config:
            target_svc = service_name.strip().lower()
            for svc in config.GetAllLoadedServices():
                if svc.Name:
                    s_name = svc.Name.strip().lower()
                    if s_name == target_svc or target_svc in s_name or s_name in target_svc:
                        grp_count = svc.PaletteCount if RevitINT > 2022 else svc.GroupCount
                        for gi in range(grp_count):
                            for bi in range(svc.GetButtonCount(gi)):
                                bt = svc.GetButton(gi, bi)
                                if bt and bt.IsAHanger and bt.Name:
                                    names.add(bt.Name.strip())
            
            if not names:
                for svc in config.GetAllLoadedServices():
                    grp_count = svc.PaletteCount if RevitINT > 2022 else svc.GroupCount
                    for gi in range(grp_count):
                        for bi in range(svc.GetButtonCount(gi)):
                            bt = svc.GetButton(gi, bi)
                            if bt and bt.IsAHanger and bt.Name:
                                names.add(bt.Name.strip())
    except:
        pass
    
    result = sorted(list(names))
    if "--- NONE ---" not in result:
        result.insert(0, "--- NONE ---")
    return result


def load_service_settings(path):
    settings = {}
    if not os.path.exists(path):
        return settings

    with open(path, 'r') as f:
        for line in f:
            line = line.strip()
            if "=" not in line: continue
            parts = line.split("=", 1)
            if len(parts) != 2: continue

            key = parts[0].strip()
            val_string = parts[1].strip()
            rules = []
            
            if ":" not in val_string:
                vals = val_string.split("|")
                if len(vals) == 3:
                    try:
                        rules.append({
                            "size": 999.0,
                            "hanger": vals[0].strip(),
                            "spacing": float(vals[1].strip()),
                            "joints": vals[2].strip().lower() == "true"
                        })
                    except: pass
            else:
                rule_strings = val_string.split("|")
                for rs in rule_strings:
                    r_parts = rs.split(":")
                    if len(r_parts) == 4:
                        try:
                            rules.append({
                                "size": float(r_parts[0].strip()),
                                "hanger": r_parts[1].strip(),
                                "spacing": float(r_parts[2].strip()),
                                "joints": r_parts[3].strip().lower() == "true"
                            })
                        except: pass
                        
            if rules:
                rules.sort(key=lambda x: x["size"])
                settings[key] = rules
    return settings


def save_service_settings(path, settings):
    with open(path, 'w') as f:
        for key in sorted(settings.keys()):
            rules = settings[key]
            rule_strings = []
            for r in rules:
                rule_strings.append("{0}:{1}:{2}:{3}".format(r["size"], r["hanger"], r["spacing"], r["joints"]))
            f.write("{0}={1}\n".format(key, " | ".join(rule_strings)))


def safe_param_string(elem, param_name):
    try:
        p = elem.LookupParameter(param_name)
        if p:
            val = p.AsString()
            if val and val.strip(): return val.strip()
            val = p.AsValueString()
            if val and val.strip(): return val.strip()
    except: pass
    return None


def get_service_name(elem):
    for p_name in ["Fabrication Service Name", "Service Name"]:
        val = safe_param_string(elem, p_name)
        if val: return val
    return "UNASSIGNED"


def collect_services_from_model(doc):
    services = set()
    try:
        elems = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_FabricationPipework)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        for elem in elems:
            svc_name = get_service_name(elem)
            if svc_name != "UNASSIGNED":
                services.add(svc_name)
    except: pass
    return sorted(list(services))


class PipeServiceForm(Window):
    def __init__(self, services, existing_settings, doc):
        self.Title = "Piping Hanger Configuration with Size Breaks ({})".format(file_name)
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.result = None
        self.inputs = {}
        self.doc = doc

        self.Width = 720
        self.Height = 800

        root = Grid()
        root.Margin = Thickness(12)
        self.Content = root

        rd0 = RowDefinition(); rd0.Height = GridLength.Auto
        rd1 = RowDefinition(); rd1.Height = GridLength(1, GridUnitType.Star)
        rd2 = RowDefinition(); rd2.Height = GridLength.Auto
        root.RowDefinitions.Add(rd0)
        root.RowDefinitions.Add(rd1)
        root.RowDefinitions.Add(rd2)

        instructions = TextBlock()
        instructions.Text = "Add size breaks for each service. Select '--- NONE ---' to disable/ignore a service or size range.\nUse '999' as the catch-all for the largest sizes."
        instructions.TextWrapping = TextWrapping.Wrap
        instructions.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(instructions, 0)
        root.Children.Add(instructions)

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        Grid.SetRow(scroll, 1)
        root.Children.Add(scroll)

        self.main_panel = StackPanel()
        scroll.Content = self.main_panel

        for service in services:
            rules = existing_settings.get(service, [])
            service_hangers = get_hangers_for_service(self.doc, service)
            default_hanger = service_hangers[0] if service_hangers else "--- NONE ---"
            
            if not rules:
                rules = [{"size": 999.0, "hanger": default_hanger, "spacing": 10.0, "joints": True}]
                
            self.build_service_block(service, rules, service_hangers, default_hanger)

        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 12, 0, 0)
        Grid.SetRow(button_panel, 2)
        root.Children.Add(button_panel)

        ok_btn = Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 90
        ok_btn.Height = 28
        ok_btn.Margin = Thickness(6, 0, 6, 0)
        ok_btn.Click += self.ok_clicked
        button_panel.Children.Add(ok_btn)

        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 90
        cancel_btn.Height = 28
        cancel_btn.Margin = Thickness(6, 0, 6, 0)
        cancel_btn.Click += self.cancel_clicked
        button_panel.Children.Add(cancel_btn)

    def build_service_block(self, service, rules, service_hangers, default_hanger):
        self.inputs[service] = []
        
        border = Border()
        border.BorderBrush = Brushes.LightGray
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(4)
        border.Margin = Thickness(0, 0, 15, 12)
        border.Padding = Thickness(8)
        border.Background = SolidColorBrush(Colors.WhiteSmoke)
        
        block_panel = StackPanel()
        border.Child = block_panel

        header_grid = Grid()
        header_grid.Margin = Thickness(0, 0, 0, 8)
        
        hc0 = ColumnDefinition(); hc0.Width = GridLength(1, GridUnitType.Star)
        hc1 = ColumnDefinition(); hc1.Width = GridLength.Auto
        header_grid.ColumnDefinitions.Add(hc0)
        header_grid.ColumnDefinitions.Add(hc1)
        
        title = TextBlock()
        title.Text = service
        title.FontWeight = FontWeights.Bold
        title.FontSize = 14
        title.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(title, 0)
        header_grid.Children.Add(title)
        
        add_btn = Button()
        add_btn.Content = "+ Add Break"
        add_btn.Width = 80
        add_btn.Height = 22
        Grid.SetColumn(add_btn, 1)
        header_grid.Children.Add(add_btn)
        
        block_panel.Children.Add(header_grid)

        col_headers = Grid()
        col_headers.Margin = Thickness(0, 0, 0, 4)
        
        cd0 = ColumnDefinition(); cd0.Width = GridLength(90)
        cd1 = ColumnDefinition(); cd1.Width = GridLength(1, GridUnitType.Star)
        cd2 = ColumnDefinition(); cd2.Width = GridLength(80)
        cd3 = ColumnDefinition(); cd3.Width = GridLength(90)
        cd4 = ColumnDefinition(); cd4.Width = GridLength(40)
        
        col_headers.ColumnDefinitions.Add(cd0)
        col_headers.ColumnDefinitions.Add(cd1)
        col_headers.ColumnDefinitions.Add(cd2)
        col_headers.ColumnDefinitions.Add(cd3)
        col_headers.ColumnDefinitions.Add(cd4)

        def make_header(text, col, align=HorizontalAlignment.Left):
            tb = TextBlock()
            tb.Text = text
            tb.FontWeight = FontWeights.Bold
            tb.FontSize = 11
            tb.HorizontalAlignment = align
            Grid.SetColumn(tb, col)
            return tb

        col_headers.Children.Add(make_header("Up To Size (\")", 0))
        col_headers.Children.Add(make_header("Hanger Type", 1))
        col_headers.Children.Add(make_header("Spacing (ft)", 2, HorizontalAlignment.Center))
        col_headers.Children.Add(make_header("Joints", 3, HorizontalAlignment.Center))
        
        block_panel.Children.Add(col_headers)

        rows_panel = StackPanel()
        block_panel.Children.Add(rows_panel)

        def make_add_handler(svc, pnl, s_hangers, d_hanger):
            def handler(sender, args):
                self.add_rule_row(svc, pnl, None, s_hangers, d_hanger)
            return handler
        add_btn.Click += make_add_handler(service, rows_panel, service_hangers, default_hanger)

        for rule in rules:
            self.add_rule_row(service, rows_panel, rule, service_hangers, default_hanger)

        self.main_panel.Children.Add(border)

    def add_rule_row(self, service, parent_panel, rule_data, service_hangers, default_hanger):
        if not rule_data:
            rule_data = {"size": 999.0, "hanger": default_hanger, "spacing": 10.0, "joints": True}
            
        row_grid = Grid()
        row_grid.Margin = Thickness(0, 2, 0, 2)
        
        cd0 = ColumnDefinition(); cd0.Width = GridLength(90)
        cd1 = ColumnDefinition(); cd1.Width = GridLength(1, GridUnitType.Star)
        cd2 = ColumnDefinition(); cd2.Width = GridLength(80)
        cd3 = ColumnDefinition(); cd3.Width = GridLength(90)
        cd4 = ColumnDefinition(); cd4.Width = GridLength(40)
        
        row_grid.ColumnDefinitions.Add(cd0)
        row_grid.ColumnDefinitions.Add(cd1)
        row_grid.ColumnDefinitions.Add(cd2)
        row_grid.ColumnDefinitions.Add(cd3)
        row_grid.ColumnDefinitions.Add(cd4)

        tb_size = TextBox()
        tb_size.Text = str(rule_data["size"])
        tb_size.Height = 22
        tb_size.Margin = Thickness(0, 0, 10, 0)
        tb_size.TextAlignment = TextAlignment.Center
        Grid.SetColumn(tb_size, 0)
        row_grid.Children.Add(tb_size)

        cb_hanger = ComboBox()
        cb_hanger.ItemsSource = Array[object](service_hangers)
        current_hanger = rule_data["hanger"]
        cb_hanger.SelectedItem = current_hanger if current_hanger in service_hangers else default_hanger
        cb_hanger.Height = 22
        cb_hanger.Margin = Thickness(0, 0, 10, 0)
        Grid.SetColumn(cb_hanger, 1)
        row_grid.Children.Add(cb_hanger)

        tb_spacing = TextBox()
        tb_spacing.Text = str(rule_data["spacing"])
        tb_spacing.Height = 22
        tb_spacing.Margin = Thickness(0, 0, 10, 0)
        tb_spacing.TextAlignment = TextAlignment.Center
        Grid.SetColumn(tb_spacing, 2)
        row_grid.Children.Add(tb_spacing)

        chk_joints = CheckBox()
        chk_joints.IsChecked = rule_data["joints"]
        chk_joints.HorizontalAlignment = HorizontalAlignment.Center
        chk_joints.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(chk_joints, 3)
        row_grid.Children.Add(chk_joints)

        del_btn = Button()
        del_btn.Content = "X"
        del_btn.Height = 22
        del_btn.Width = 22
        del_btn.Foreground = Brushes.Red
        del_btn.FontWeight = FontWeights.Bold
        Grid.SetColumn(del_btn, 4)
        row_grid.Children.Add(del_btn)

        rule_controls = {
            "size": tb_size, 
            "hanger": cb_hanger, 
            "spacing": tb_spacing, 
            "joints": chk_joints,
            "ui_grid": row_grid
        }
        
        self.inputs[service].append(rule_controls)

        def make_del_handler(svc, pnl, r_dict):
            def handler(sender, args):
                if len(self.inputs[svc]) > 1:
                    pnl.Children.Remove(r_dict["ui_grid"])
                    self.inputs[svc].remove(r_dict)
            return handler

        del_btn.Click += make_del_handler(service, parent_panel, rule_controls)
        parent_panel.Children.Add(row_grid)

    def ok_clicked(self, sender, args):
        values = {}
        for service, rules_list in self.inputs.items():
            parsed_rules = []
            for controls in rules_list:
                try:
                    size_text = controls["size"].Text.strip().upper()
                    if size_text in ["ANY", "MAX", "ALL"]:
                        size_val = 999.0
                    else:
                        size_val = float(size_text)
                        
                    spacing_val = float(controls["spacing"].Text)
                    selected_hanger = controls["hanger"].SelectedItem
                    hanger_val = str(selected_hanger).strip() if selected_hanger else "--- NONE ---"
                    
                    parsed_rules.append({
                        "size": size_val,
                        "hanger": hanger_val,
                        "spacing": spacing_val,
                        "joints": bool(controls["joints"].IsChecked)
                    })
                except:
                    controls["size"].Focus()
                    controls["size"].SelectAll()
                    return

            parsed_rules.sort(key=lambda x: x["size"])
            values[service] = parsed_rules

        self.result = values
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()

try:
    services_in_model = collect_services_from_model(doc)
    saved_settings = load_service_settings(FILEPATH)

    if not services_in_model:
        TaskDialog.Show("Pipe Hangers", "No fabrication piping services found in the current model.")
    else:
        form = PipeServiceForm(services_in_model, saved_settings, doc)
        if form.ShowDialog():
            if form.result:
                save_service_settings(FILEPATH, form.result)
                
except Exception as e:
    TaskDialog.Show("Error Generating Config", traceback.format_exc())