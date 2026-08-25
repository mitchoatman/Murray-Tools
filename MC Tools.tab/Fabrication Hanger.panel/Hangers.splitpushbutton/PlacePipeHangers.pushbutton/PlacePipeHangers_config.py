# -*- coding: UTF-8 -*-
import os
import clr
import traceback
import System
import re
import Autodesk
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

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, FabricationConfiguration, Transaction, ElementId
from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, Entity, FieldBuilder, AccessLevel
from Autodesk.Revit.UI import TaskDialog

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
RevitINT = float(app.VersionNumber)

# Unique GUID for our Extensible Storage Schema
SCHEMA_GUID = System.Guid("7B3F8A12-4C9E-4D21-8F6B-1E9A3C5D7F8E")
DATA_STORAGE_NAME = "PipeHangerConfigData"


def get_or_create_schema():
    schema = Schema.Lookup(SCHEMA_GUID)
    if schema is not None:
        return schema

    builder = SchemaBuilder(SCHEMA_GUID)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.SetVendorId("BIMTools")
    builder.SetSchemaName("PipeHangerConfigSchema")
    
    builder.AddSimpleField("ConfigPayload", System.String)
    return builder.Finish()


def normalize_service_name(name):
    if not name:
        return ""
    return re.sub(r"\s+", " ", name).strip().lower()


def try_get_int_id(value):
    try:
        return value.IntegerValue
    except:
        try:
            return int(value)
        except:
            return None


def get_element_service_id(elem):
    try:
        if hasattr(elem, "ServiceId"):
            sid = try_get_int_id(elem.ServiceId)
            if sid is not None and sid >= 0:
                return sid
    except:
        pass
    return None


def get_loaded_service_id(svc):
    for attr_name in ["ServiceId", "Id"]:
        try:
            if hasattr(svc, attr_name):
                sid = try_get_int_id(getattr(svc, attr_name))
                if sid is not None and sid >= 0:
                    return sid
        except:
            pass
    return None


def get_hangers_for_service(doc, service_id, service_name):
    """Retrieves hangers only from the matching fabrication service."""
    names = set()
    try:
        config = FabricationConfiguration.GetFabricationConfiguration(doc)
        if config:
            target_name = normalize_service_name(service_name)

            for svc in config.GetAllLoadedServices():
                is_match = False

                loaded_id = get_loaded_service_id(svc)
                if service_id is not None and loaded_id is not None:
                    is_match = (loaded_id == service_id)
                else:
                    svc_name = normalize_service_name(svc.Name)
                    is_match = (svc_name == target_name)

                if not is_match:
                    continue

                grp_count = svc.PaletteCount if RevitINT > 2022 else svc.GroupCount
                for gi in range(grp_count):
                    for bi in range(svc.GetButtonCount(gi)):
                        bt = svc.GetButton(gi, bi)
                        if bt and bt.IsAHanger and bt.Name:
                            names.add(bt.Name.strip())
                break
    except:
        pass

    result = sorted(list(names))
    if "--- NONE ---" not in result:
        result.insert(0, "--- NONE ---")
    return result


def load_service_settings(doc):
    settings = {}
    try:
        schema = Schema.Lookup(SCHEMA_GUID)
        if not schema:
            return settings

        collector = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.ExtensibleStorage.DataStorage)
        target_ds = None
        for ds in collector:
            if ds.Name == DATA_STORAGE_NAME:
                target_ds = ds
                break

        if not target_ds:
            return settings

        entity = target_ds.GetEntity(schema)
        if not entity.IsValid():
            return settings

        val_string = entity.Get[System.String]("ConfigPayload")
        if not val_string:
            return settings

        lines = val_string.split("\n")
        for line in lines:
            line = line.strip()
            if "=" not in line:
                continue
            parts = line.split("=", 1)
            if len(parts) != 2:
                continue

            key = parts[0].strip()
            rule_block = parts[1].strip()
            rules = []

            rule_strings = rule_block.split("|")
            for rs in rule_strings:
                r_parts = rs.split(":")
                if len(r_parts) == 5:
                    try:
                        rules.append({
                            "size": float(r_parts[0].strip()),
                            "hanger": r_parts[1].strip(),
                            "spacing": float(r_parts[2].strip()),
                            "dist_from_end": float(r_parts[3].strip()),
                            "joints": r_parts[4].strip().lower() == "true"
                        })
                    except:
                        pass

            if rules:
                rules.sort(key=lambda x: x["size"])
                settings[key] = rules
    except:
        pass
    return settings


def save_service_settings(doc, settings):
    lines = []
    for key in sorted(settings.keys()):
        rules = settings[key]
        rule_strings = []
        for r in rules:
            rule_strings.append("{0}:{1}:{2}:{3}:{4}".format(
                r["size"],
                r["hanger"],
                r["spacing"],
                r["dist_from_end"],
                r["joints"]
            ))
        lines.append("{0}={1}".format(key, " | ".join(rule_strings)))

    combined_text = "\n".join(lines)
    schema = get_or_create_schema()

    t = Transaction(doc, "Save Pipe Hanger Configurations")
    t.Start()
    try:
        collector = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.ExtensibleStorage.DataStorage)
        target_ds = None
        for ds in collector:
            if ds.Name == DATA_STORAGE_NAME:
                target_ds = ds
                break

        if not target_ds:
            target_ds = Autodesk.Revit.DB.ExtensibleStorage.DataStorage.Create(doc)
            target_ds.Name = DATA_STORAGE_NAME

        entity = Entity(schema)
        entity.Set("ConfigPayload", combined_text)
        target_ds.SetEntity(entity)
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error Saving Config", str(e))


def safe_param_string(elem, param_name):
    try:
        p = elem.LookupParameter(param_name)
        if p:
            val = p.AsString()
            if val and val.strip():
                return val.strip()
            val = p.AsValueString()
            if val and val.strip():
                return val.strip()
    except:
        pass
    return None


def get_service_name(elem):
    for p_name in ["Fabrication Service Name", "Service Name"]:
        val = safe_param_string(elem, p_name)
        if val:
            return val
    return "UNASSIGNED"


def collect_services_from_model(doc):
    services = {}
    try:
        elems = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_FabricationPipework)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        for elem in elems:
            svc_name = get_service_name(elem)
            if svc_name == "UNASSIGNED":
                continue

            svc_id = get_element_service_id(elem)

            if svc_name not in services:
                services[svc_name] = {
                    "name": svc_name,
                    "id": svc_id
                }
            elif services[svc_name]["id"] is None and svc_id is not None:
                services[svc_name]["id"] = svc_id
    except:
        pass

    return [services[k] for k in sorted(services.keys())]


class PipeServiceForm(Window):
    def __init__(self, services, existing_settings, doc):
        file_path = doc.PathName
        file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)
        if not file_name:
            file_name = doc.Title

        self.base_title = "Piping Hanger Configuration ({})".format(file_name)
        self.Title = self.base_title
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.result = None
        self.inputs = {}
        self.panels = {}
        self.service_hangers_map = {}
        self.copied_rules = None
        self.last_copied_title_block = None
        self.last_copied_service_name = None
        self.doc = doc
        self.service_records = services
        self.services = [s["name"] for s in services]

        self.Width = 760
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

        for service_record in services:
            service = service_record["name"]
            service_id = service_record["id"]

            rules = existing_settings.get(service, [])
            service_hangers = get_hangers_for_service(self.doc, service_id, service)
            default_hanger = service_hangers[0] if service_hangers else "--- NONE ---"
            
            self.service_hangers_map[service] = (service_hangers, default_hanger)

            if not rules:
                rules = [{
                    "size": 999.0,
                    "hanger": default_hanger,
                    "spacing": 8.0,
                    "dist_from_end": 1.0,
                    "joints": False
                }]

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

        defaults_btn = Button()
        defaults_btn.Content = "Defaults"
        defaults_btn.Width = 90
        defaults_btn.Height = 28
        defaults_btn.Margin = Thickness(6, 0, 6, 0)
        defaults_btn.Click += self.defaults_clicked
        button_panel.Children.Add(defaults_btn)

        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 90
        cancel_btn.Height = 28
        cancel_btn.Margin = Thickness(6, 0, 6, 0)
        cancel_btn.Click += self.cancel_clicked
        button_panel.Children.Add(cancel_btn)

    def clear_copied_indicator(self):
        if self.last_copied_title_block is not None and self.last_copied_service_name is not None:
            self.last_copied_title_block.Text = self.last_copied_service_name

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
        hc2 = ColumnDefinition(); hc2.Width = GridLength.Auto
        hc3 = ColumnDefinition(); hc3.Width = GridLength.Auto
        header_grid.ColumnDefinitions.Add(hc0)
        header_grid.ColumnDefinitions.Add(hc1)
        header_grid.ColumnDefinitions.Add(hc2)
        header_grid.ColumnDefinitions.Add(hc3)

        title = TextBlock()
        title.Text = service
        title.FontWeight = FontWeights.Bold
        title.FontSize = 14
        title.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(title, 0)
        header_grid.Children.Add(title)

        copy_btn = Button()
        copy_btn.Content = "Copy"
        copy_btn.Width = 55
        copy_btn.Height = 22
        copy_btn.Margin = Thickness(0, 0, 4, 0)
        Grid.SetColumn(copy_btn, 1)
        header_grid.Children.Add(copy_btn)

        paste_btn = Button()
        paste_btn.Content = "Paste"
        paste_btn.Width = 55
        paste_btn.Height = 22
        paste_btn.Margin = Thickness(0, 0, 4, 0)
        Grid.SetColumn(paste_btn, 2)
        header_grid.Children.Add(paste_btn)

        add_btn = Button()
        add_btn.Content = "+ Add Break"
        add_btn.Width = 80
        add_btn.Height = 22
        Grid.SetColumn(add_btn, 3)
        header_grid.Children.Add(add_btn)

        block_panel.Children.Add(header_grid)

        col_headers = Grid()
        col_headers.Margin = Thickness(0, 0, 0, 4)

        cd0 = ColumnDefinition(); cd0.Width = GridLength(90)
        cd1 = ColumnDefinition(); cd1.Width = GridLength(1, GridUnitType.Star)
        cd2 = ColumnDefinition(); cd2.Width = GridLength(80)
        cd3 = ColumnDefinition(); cd3.Width = GridLength(95)
        cd4 = ColumnDefinition(); cd4.Width = GridLength(55)
        cd5 = ColumnDefinition(); cd5.Width = GridLength(40)

        col_headers.ColumnDefinitions.Add(cd0)
        col_headers.ColumnDefinitions.Add(cd1)
        col_headers.ColumnDefinitions.Add(cd2)
        col_headers.ColumnDefinitions.Add(cd3)
        col_headers.ColumnDefinitions.Add(cd4)
        col_headers.ColumnDefinitions.Add(cd5)

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
        col_headers.Children.Add(make_header("Dist End (ft)", 3, HorizontalAlignment.Center))
        col_headers.Children.Add(make_header("Joints", 4, HorizontalAlignment.Center))

        block_panel.Children.Add(col_headers)

        rows_panel = StackPanel()
        block_panel.Children.Add(rows_panel)
        self.panels[service] = rows_panel

        def make_add_handler(svc, pnl, s_hangers, d_hanger):
            def handler(sender, args):
                self.add_rule_row(svc, pnl, None, s_hangers, d_hanger)
            return handler
        add_btn.Click += make_add_handler(service, rows_panel, service_hangers, default_hanger)

        def make_copy_handler(svc, t_block):
            def handler(sender, args):
                self.copy_service_settings(svc, t_block)
            return handler
        copy_btn.Click += make_copy_handler(service, title)

        def make_paste_handler(svc, pnl, s_hangers, d_hanger):
            def handler(sender, args):
                self.paste_service_settings(svc, pnl, s_hangers, d_hanger)
            return handler
        paste_btn.Click += make_paste_handler(service, rows_panel, service_hangers, default_hanger)

        for rule in rules:
            self.add_rule_row(service, rows_panel, rule, service_hangers, default_hanger)

        self.main_panel.Children.Add(border)

    def copy_service_settings(self, service_name, title_block):
        current_rules = []
        for controls in self.inputs.get(service_name, []):
            try:
                size_text = controls["size"].Text.strip().upper()
                size_val = 999.0 if size_text in ["ANY", "MAX", "ALL"] else float(size_text)
                spacing_val = float(controls["spacing"].Text)
                dist_end_val = float(controls["dist_from_end"].Text)
                selected_hanger = controls["hanger"].SelectedItem
                hanger_val = str(selected_hanger).strip() if selected_hanger else "--- NONE ---"

                current_rules.append({
                    "size": size_val,
                    "hanger": hanger_val,
                    "spacing": spacing_val,
                    "dist_from_end": dist_end_val,
                    "joints": bool(controls["joints"].IsChecked)
                })
            except:
                pass
        
        if current_rules:
            self.clear_copied_indicator()
            self.copied_rules = current_rules
            title_block.Text = "{} [COPIED]".format(service_name)
            self.Title = "Piping Hanger Configuration — [Copied: {}]".format(service_name)
            self.last_copied_title_block = title_block
            self.last_copied_service_name = service_name

    def paste_service_settings(self, service_name, parent_panel, service_hangers, default_hanger):
        if not self.copied_rules:
            TaskDialog.Show("Hanger Config", "No settings have been copied yet. Click 'Copy' on a service first.")
            return

        parent_panel.Children.Clear()
        self.inputs[service_name] = []

        for rule in self.copied_rules:
            self.add_rule_row(service_name, parent_panel, rule, service_hangers, default_hanger)

    def add_rule_row(self, service, parent_panel, rule_data, service_hangers, default_hanger):
        if not rule_data:
            rule_data = {
                "size": 999.0,
                "hanger": default_hanger,
                "spacing": 8.0,
                "dist_from_end": 1.0,
                "joints": False
            }

        row_grid = Grid()
        row_grid.Margin = Thickness(0, 2, 0, 2)

        cd0 = ColumnDefinition(); cd0.Width = GridLength(90)
        cd1 = ColumnDefinition(); cd1.Width = GridLength(1, GridUnitType.Star)
        cd2 = ColumnDefinition(); cd2.Width = GridLength(80)
        cd3 = ColumnDefinition(); cd3.Width = GridLength(95)
        cd4 = ColumnDefinition(); cd4.Width = GridLength(55)
        cd5 = ColumnDefinition(); cd5.Width = GridLength(40)

        row_grid.ColumnDefinitions.Add(cd0)
        row_grid.ColumnDefinitions.Add(cd1)
        row_grid.ColumnDefinitions.Add(cd2)
        row_grid.ColumnDefinitions.Add(cd3)
        row_grid.ColumnDefinitions.Add(cd4)
        row_grid.ColumnDefinitions.Add(cd5)

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

        tb_dist_end = TextBox()
        tb_dist_end.Text = str(rule_data.get("dist_from_end", 1.0))
        tb_dist_end.Height = 22
        tb_dist_end.Margin = Thickness(0, 0, 10, 0)
        tb_dist_end.TextAlignment = TextAlignment.Center
        Grid.SetColumn(tb_dist_end, 3)
        row_grid.Children.Add(tb_dist_end)

        chk_joints = CheckBox()
        chk_joints.IsChecked = rule_data["joints"]
        chk_joints.HorizontalAlignment = HorizontalAlignment.Center
        chk_joints.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(chk_joints, 4)
        row_grid.Children.Add(chk_joints)

        del_btn = Button()
        del_btn.Content = "X"
        del_btn.Height = 22
        del_btn.Width = 22
        del_btn.Foreground = Brushes.Red
        del_btn.FontWeight = FontWeights.Bold
        Grid.SetColumn(del_btn, 5)
        row_grid.Children.Add(del_btn)

        rule_controls = {
            "size": tb_size,
            "hanger": cb_hanger,
            "spacing": tb_spacing,
            "dist_from_end": tb_dist_end,
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

    def defaults_clicked(self, sender, args):
        for service in self.services:
            pnl = self.panels.get(service)
            if not pnl:
                continue
            
            pnl.Children.Clear()
            self.inputs[service] = []

            service_hangers, default_hanger = self.service_hangers_map.get(service, ([], "--- NONE ---"))
            
            default_rule = {
                "size": 999.0,
                "hanger": default_hanger,
                "spacing": 8.0,
                "dist_from_end": 1.0,
                "joints": False
            }
            self.add_rule_row(service, pnl, default_rule, service_hangers, default_hanger)

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
                    dist_end_val = float(controls["dist_from_end"].Text)
                    selected_hanger = controls["hanger"].SelectedItem
                    hanger_val = str(selected_hanger).strip() if selected_hanger else "--- NONE ---"

                    parsed_rules.append({
                        "size": size_val,
                        "hanger": hanger_val,
                        "spacing": spacing_val,
                        "dist_from_end": dist_end_val,
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
    saved_settings = load_service_settings(doc)

    if not services_in_model:
        TaskDialog.Show("Pipe Hangers", "No fabrication piping services found in the current model.")
    else:
        form = PipeServiceForm(services_in_model, saved_settings, doc)
        if form.ShowDialog():
            if form.result:
                save_service_settings(doc, form.result)

except Exception as e:
    TaskDialog.Show("Error Generating Config", traceback.format_exc())