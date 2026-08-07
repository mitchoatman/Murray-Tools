# -*- coding: utf-8 -*-
import os
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Array
from System.Windows import (
    Window, Thickness, WindowStartupLocation, ResizeMode,
    HorizontalAlignment, VerticalAlignment, GridLength, SizeToContent
)
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, Label, TextBox,
    Button, ComboBox, CheckBox, StackPanel
)
from System.Windows.Media import FontFamily

from Autodesk.Revit.DB import FabricationConfiguration

doc = __revit__.ActiveUIDocument.Document
app = doc.Application
RevitINT = float(app.VersionNumber)

FOLDER = r"C:\Temp"
FILEPATH = os.path.join(FOLDER, "Ribbon_Duct-Hanger-Config.txt")

DEFAULTS = {
    "ROUND_HANGER": "",
    "RECT_HANGER": "",
    "END_DIST_IN": "12",
    "ROUND_MAX_SPACING_FT": "8",
    "RECT_MAX_SPACING_FT": "8",
    "ATTACH_TO_STRUCTURE": "True",
}

CONFIG_KEYS = [
    "ROUND_HANGER",
    "RECT_HANGER",
    "END_DIST_IN",
    "ROUND_MAX_SPACING_FT",
    "RECT_MAX_SPACING_FT",
    "ATTACH_TO_STRUCTURE",
]


def ensure_folder():
    if not os.path.exists(FOLDER):
        os.makedirs(FOLDER)


def load_config(path):
    data = dict(DEFAULTS)

    if not os.path.exists(path):
        return data

    with open(path, "r") as f:
        for line in f:
            line = line.strip()
            if "=" not in line:
                continue
            key, val = line.split("=", 1)
            key = key.strip()
            val = val.strip()
            if key:
                data[key] = val

    return data


def save_config(path, data):
    with open(path, "w") as f:
        for key in CONFIG_KEYS:
            f.write("{0}={1}\n".format(key, data.get(key, "")))


def collect_hanger_names(doc):
    names = set()

    config = FabricationConfiguration.GetFabricationConfiguration(doc)
    services = config.GetAllLoadedServices()

    for svc in services:
        try:
            grp_count = svc.PaletteCount if RevitINT > 2022 else svc.GroupCount
            for gi in range(grp_count):
                for bi in range(svc.GetButtonCount(gi)):
                    bt = svc.GetButton(gi, bi)
                    if bt.IsAHanger:
                        try:
                            if bt.Name and bt.Name.strip():
                                names.add(bt.Name.strip())
                        except:
                            pass
        except:
            pass

    result = sorted(list(names))
    if not result:
        result = [""]
    return result


class DuctHangerConfigForm(Window):
    def __init__(self, hanger_names, saved):
        self.Title = "Duct Hanger Configuration"
        self.Width = 520
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.SizeToContent = SizeToContent.Height  
        
        self.result = None

        root = Grid()
        root.Margin = Thickness(12)
        self.Content = root

        for _ in range(7):
            row_def = RowDefinition()
            row_def.Height = GridLength.Auto
            root.RowDefinitions.Add(row_def)

        root.ColumnDefinitions.Add(ColumnDefinition())
        root.ColumnDefinitions.Add(ColumnDefinition())

        def make_label(text):
            lbl = Label()
            lbl.Content = text
            lbl.FontFamily = FontFamily("Arial")
            lbl.FontSize = 12
            lbl.Margin = Thickness(0, 4, 0, 4) 
            return lbl

        def make_textbox(text):
            tb = TextBox()
            tb.Text = text
            tb.Width = 180
            tb.Height = 24
            tb.FontFamily = FontFamily("Arial")
            tb.FontSize = 12
            tb.Margin = Thickness(0, 4, 0, 4)
            return tb

        def add_control(row, left_text, control):
            lbl = make_label(left_text)
            Grid.SetRow(lbl, row)
            Grid.SetColumn(lbl, 0)
            root.Children.Add(lbl)

            Grid.SetRow(control, row)
            Grid.SetColumn(control, 1)
            root.Children.Add(control)

        self.cb_round_hanger = ComboBox()
        self.cb_round_hanger.Width = 220
        self.cb_round_hanger.Height = 24
        self.cb_round_hanger.Margin = Thickness(0, 4, 0, 4)
        self.cb_round_hanger.ItemsSource = Array[object](hanger_names)
        self.cb_round_hanger.SelectedItem = saved["ROUND_HANGER"] if saved["ROUND_HANGER"] in hanger_names else hanger_names[0]
        add_control(0, "Round Duct Hanger:", self.cb_round_hanger)

        self.cb_rect_hanger = ComboBox()
        self.cb_rect_hanger.Width = 220
        self.cb_rect_hanger.Height = 24
        self.cb_rect_hanger.Margin = Thickness(0, 4, 0, 4)
        self.cb_rect_hanger.ItemsSource = Array[object](hanger_names)
        self.cb_rect_hanger.SelectedItem = saved["RECT_HANGER"] if saved["RECT_HANGER"] in hanger_names else hanger_names[0]
        add_control(1, "Rectangular Duct Hanger:", self.cb_rect_hanger)

        self.tb_end_dist = make_textbox(saved["END_DIST_IN"])
        add_control(2, "Distance From End (in):", self.tb_end_dist)

        self.tb_round_spacing = make_textbox(saved["ROUND_MAX_SPACING_FT"])
        add_control(3, "Round Max Spacing (ft):", self.tb_round_spacing)

        self.tb_rect_spacing = make_textbox(saved["RECT_MAX_SPACING_FT"])
        add_control(4, "Rectangular Max Spacing (ft):", self.tb_rect_spacing)

        self.chk_atos = CheckBox()
        self.chk_atos.Content = "Attach to Structure"
        self.chk_atos.IsChecked = str(saved["ATTACH_TO_STRUCTURE"]).lower() == "true"
        self.chk_atos.FontFamily = FontFamily("Arial")
        self.chk_atos.FontSize = 12
        self.chk_atos.Margin = Thickness(0, 4, 0, 4)
        add_control(5, "", self.chk_atos)

        btn_panel = StackPanel()
        btn_panel.Orientation = 0
        btn_panel.HorizontalAlignment = HorizontalAlignment.Center
        btn_panel.Margin = Thickness(0, 16, 0, 8)

        ok_btn = Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 90
        ok_btn.Height = 28
        ok_btn.Margin = Thickness(6)
        ok_btn.Click += self.ok_clicked
        btn_panel.Children.Add(ok_btn)

        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 90
        cancel_btn.Height = 28
        cancel_btn.Margin = Thickness(6)
        cancel_btn.Click += self.cancel_clicked
        btn_panel.Children.Add(cancel_btn)

        Grid.SetRow(btn_panel, 6)
        Grid.SetColumnSpan(btn_panel, 2)
        root.Children.Add(btn_panel)

    def validate_float(self, tb):
        try:
            float(tb.Text)
            return True
        except:
            tb.Focus()
            tb.SelectAll()
            return False

    def ok_clicked(self, sender, args):
        for tb in [
            self.tb_end_dist,
            self.tb_round_spacing,
            self.tb_rect_spacing,
        ]:
            if not self.validate_float(tb):
                return

        self.result = {
            "ROUND_HANGER": self.cb_round_hanger.SelectedItem or "",
            "RECT_HANGER": self.cb_rect_hanger.SelectedItem or "",
            "END_DIST_IN": self.tb_end_dist.Text,
            "ROUND_MAX_SPACING_FT": self.tb_round_spacing.Text,
            "RECT_MAX_SPACING_FT": self.tb_rect_spacing.Text,
            "ATTACH_TO_STRUCTURE": str(bool(self.chk_atos.IsChecked)),
        }

        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


ensure_folder()
saved = load_config(FILEPATH)
hanger_names = collect_hanger_names(doc)

form = DuctHangerConfigForm(hanger_names, saved)
if form.ShowDialog():
    if form.result:
        save_config(FILEPATH, form.result)