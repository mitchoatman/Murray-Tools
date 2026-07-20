# -*- coding: utf-8 -*-
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, HorizontalAlignment
from System.Windows.Controls import StackPanel, Label, TextBox, Button, Orientation
from System.Windows.Media import FontFamily

import re

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def parse_inches_part(text):
    text = text.strip().replace('"', '')
    if not text:
        return 0.0

    total_inches = 0.0
    parts = text.split()

    for part in parts:
        if '/' in part:
            num, den = part.split('/')
            total_inches += float(num) / float(den)
        else:
            total_inches += float(part)

    return total_inches


def parse_length_to_feet(input_str):
    """
    Accepts:
        15         -> 15 feet
        10 6       -> 10 feet 6 inches
        10 6 1/2   -> 10 feet 6 1/2 inches
        3.5        -> 3.5 feet
        3"         -> 3 inches
        6 1/2"     -> 6.5 inches
        1'         -> 1 foot
        1' 6"      -> 1.5 feet
        1'-6 1/2"  -> 1.541666...
    Rules:
        - If input contains ' , parse as feet + optional inches
        - If input contains " , parse as inches
        - If input has multiple space-separated number blocks and no unit marks,
          first block = feet, remaining block(s) = inches/fraction
        - Plain single number = feet
    Returns feet as float.
    """
    s = input_str.strip()
    if not s:
        raise ValueError("Empty input.")

    s = s.replace("’", "'").replace("′", "'")
    s = s.replace("”", '"').replace("″", '"')
    s = re.sub(r"\s*-\s*", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    # Explicit feet mark
    if "'" in s:
        match = re.match(r"^\s*(-?\d+(?:\.\d+)?)\s*'\s*(.*)$", s)
        if not match:
            raise ValueError("Invalid feet/inches format.")

        feet = float(match.group(1))
        inch_text = match.group(2).strip()
        inches = parse_inches_part(inch_text) if inch_text else 0.0
        return feet + (inches / 12.0)

    # Explicit inches mark
    if '"' in s:
        inches = parse_inches_part(s)
        return inches / 12.0

    # No unit marks:
    parts = s.split()

    # One plain number = feet
    if len(parts) == 1:
        return float(parts[0])

    # Multiple blocks = feet + inches/fraction
    feet = float(parts[0])
    inches = parse_inches_part(" ".join(parts[1:]))
    return feet + (inches / 12.0)


def set_parameter_by_name(element, parameter_name, value):
    param = element.LookupParameter(parameter_name)
    if param and not param.IsReadOnly:
        param.Set(value)
        return True
    return False


def get_parameter_value_by_name(element, parameter_name):
    param = element.LookupParameter(parameter_name)
    if param:
        return param.AsDouble()
    return None


class PipeAccessoryFilter(ISelectionFilter):
    def AllowElement(self, e):
        return e.Category and e.Category.Name == "Pipe Accessories"

    def AllowReference(self, ref, point):
        return False


class RodElevationDialog(Window):
    def __init__(self, default_value='3"'):
        self.Title = "Rod Elevation"
        self.Width = 400
        self.Height = 210
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize

        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        stack.Margin = Thickness(15)

        def make_label(text):
            lbl = Label()
            lbl.Content = text
            lbl.FontSize = 13
            lbl.FontFamily = FontFamily("Segoe UI")
            return lbl

        stack.Children.Add(make_label("Enter rod elevation / deck elevation:"))
        stack.Children.Add(make_label("Examples: 3 = 3'-0\", 3\", 6 1/2\", 1' 6\", 1'-6 1/2\""))

        self.txt_value = TextBox()
        self.txt_value.Text = default_value
        self.txt_value.Width = 220
        self.txt_value.HorizontalAlignment = HorizontalAlignment.Left
        self.txt_value.Margin = Thickness(0, 5, 0, 15)
        stack.Children.Add(self.txt_value)

        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 5, 0, 0)

        self.btn_ok = Button()
        self.btn_ok.Content = "OK"
        self.btn_ok.Width = 80
        self.btn_ok.Height = 30
        self.btn_ok.Margin = Thickness(0, 0, 10, 0)
        self.btn_ok.IsDefault = True
        self.btn_ok.Click += self.on_ok_clicked
        button_panel.Children.Add(self.btn_ok)

        self.btn_cancel = Button()
        self.btn_cancel.Content = "Cancel"
        self.btn_cancel.Width = 80
        self.btn_cancel.Height = 30
        self.btn_cancel.IsCancel = True
        button_panel.Children.Add(self.btn_cancel)

        stack.Children.Add(button_panel)

        self.Content = stack
        self.Loaded += self.on_loaded

    def on_loaded(self, sender, event):
        self.txt_value.Focus()
        self.txt_value.SelectAll()

    def on_ok_clicked(self, sender, event):
        self.DialogResult = True
        self.Close()


try:
    picked = uidoc.Selection.PickObjects(
        ObjectType.Element,
        PipeAccessoryFilter(),
        "Select Stiffy"
    )
    hangers = [doc.GetElement(x.ElementId) for x in picked]
except:
    hangers = []

if hangers:
    form = RodElevationDialog('15')

    if form.ShowDialog():
        try:
            target_elevation = parse_length_to_feet(form.txt_value.Text)
        except Exception:
            TaskDialog.Show("Input Error", "Invalid format.\nExamples:\n3\"\n6 1/2\"\n1' 6\"\n1'-6 1/2\"")
        else:
            t = Transaction(doc, "Set Rod Elevation")
            t.Start()

            for hanger in hangers:
                hgr_elev = get_parameter_value_by_name(hanger, "Elevation from Level")
                if hgr_elev is None:
                    continue

                new_rod_len = target_elevation - hgr_elev

                set_parameter_by_name(hanger, "Rod Elevation", new_rod_len)
                set_parameter_by_name(hanger, "DIM D", new_rod_len)

            t.Commit()