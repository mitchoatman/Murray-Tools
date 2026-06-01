# -*- coding: utf-8 -*-
import Autodesk
import clr
import os
import sys

# Revit / .NET references
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

import System
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import View, ViewType, FilteredElementCollector
from System.Windows.Forms import FolderBrowserDialog, DialogResult

from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, ResizeMode
from System.Windows import GridLength, GridUnitType
from System.Windows.Controls import (
    Label, TextBox, Button, Grid, StackPanel, ScrollViewer, CheckBox,
    RowDefinition, ColumnDefinition, Orientation, ScrollBarVisibility
)
from System.Windows.Media import FontFamily

# Active Revit document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


def sanitize_filename(name):
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for ch in invalid_chars:
        name = name.replace(ch, '_')
    return name


def is_valid_export_view(view):
    if not isinstance(view, View):
        return False
    if view.IsTemplate:
        return False
    if view.ViewType not in [ViewType.FloorPlan, ViewType.CeilingPlan]:
        return False
    return True


class ViewSelectionDialog(Window):
    def __init__(self, views, title, label_text):
        self.all_views = sorted(views, key=lambda x: x.Name)
        self.selected_ids = set()
        self.checkboxes = []
        self.check_all_state = False
        self.dialog_title = title
        self.label_text = label_text
        self.selected_views = []
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = self.dialog_title
        self.Width = 450
        self.Height = 500
        self.MinWidth = 450
        self.MinHeight = 500
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(8)

        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.ColumnDefinitions.Add(ColumnDefinition())

        self.label = Label(Content=self.label_text)
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 14
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        self.search_box = TextBox(Height=24, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        self.checkbox_panel = StackPanel(Orientation=Orientation.Vertical)
        scroll = ScrollViewer()
        scroll.Content = self.checkbox_panel
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Margin = Thickness(0, 4, 0, 4)
        Grid.SetRow(scroll, 2)
        grid.Children.Add(scroll)

        self.update_checkboxes(self.all_views)

        button_panel = StackPanel(Orientation=Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center)
        button_panel.Margin = Thickness(0, 8, 0, 0)

        self.select_button = Button(Content="Next", Width=80, Height=28, Margin=Thickness(8, 0, 8, 0))
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", Width=90, Height=28, Margin=Thickness(8, 0, 8, 0))
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        self.cancel_button = Button(Content="Cancel", Width=80, Height=28, Margin=Thickness(8, 0, 8, 0))
        self.cancel_button.Click += self.cancel_clicked
        button_panel.Children.Add(self.cancel_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        self.Content = grid

    def get_display_name(self, view):
        return "{} ({})".format(view.Name, view.ViewType)

    def update_checkboxes(self, views):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []

        for view in views:
            cb = CheckBox()
            cb.Content = self.get_display_name(view)
            cb.Tag = view.Id.IntegerValue
            cb.Margin = Thickness(2)
            cb.IsChecked = (view.Id.IntegerValue in self.selected_ids)
            cb.Click += self.checkbox_clicked
            self.checkbox_panel.Children.Add(cb)
            self.checkboxes.append(cb)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower().strip()
        if not search_text:
            filtered = self.all_views
        else:
            filtered = [v for v in self.all_views if search_text in v.Name.lower()]
        self.update_checkboxes(filtered)

    def checkbox_clicked(self, sender, args):
        view_id = sender.Tag
        if sender.IsChecked:
            self.selected_ids.add(view_id)
        else:
            if view_id in self.selected_ids:
                self.selected_ids.remove(view_id)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
            if self.check_all_state:
                self.selected_ids.add(cb.Tag)
            else:
                if cb.Tag in self.selected_ids:
                    self.selected_ids.remove(cb.Tag)

    def select_clicked(self, sender, args):
        self.selected_views = [v for v in self.all_views if v.Id.IntegerValue in self.selected_ids]
        if not self.selected_views:
            TaskDialog.Show("Export DWG", "No views selected.")
            return
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


class ExportFolderAndOptionsDialog(Window):
    def __init__(self):
        self.selected_folder = None
        self.export_settings = None
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "DWG Export Options"
        self.Width = 520
        self.Height = 300
        self.MinWidth = 520
        self.MinHeight = 300
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(10)

        for _ in range(8):
            grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

        grid.ColumnDefinitions.Add(ColumnDefinition())
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength.Auto))

        lbl = Label(Content="Select export folder and DWG settings")
        lbl.FontFamily = FontFamily("Arial")
        lbl.FontSize = 14
        Grid.SetRow(lbl, 0)
        Grid.SetColumnSpan(lbl, 2)
        grid.Children.Add(lbl)

        folder_lbl = Label(Content="Output Folder")
        folder_lbl.FontFamily = FontFamily("Arial")
        Grid.SetRow(folder_lbl, 1)
        Grid.SetColumnSpan(folder_lbl, 2)
        grid.Children.Add(folder_lbl)

        default_folder = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        self.folder_box = TextBox(Text=default_folder, Height=24, FontFamily=FontFamily("Arial"))
        Grid.SetRow(self.folder_box, 2)
        Grid.SetColumn(self.folder_box, 0)
        grid.Children.Add(self.folder_box)

        browse_btn = Button(Content="Browse...", Width=90, Height=24, Margin=Thickness(8, 0, 0, 0))
        browse_btn.Click += self.browse_clicked
        Grid.SetRow(browse_btn, 2)
        Grid.SetColumn(browse_btn, 1)
        grid.Children.Add(browse_btn)

        options_lbl = Label(Content="DWG Export Settings")
        options_lbl.FontFamily = FontFamily("Arial")
        options_lbl.Margin = Thickness(0, 10, 0, 0)
        Grid.SetRow(options_lbl, 3)
        Grid.SetColumnSpan(options_lbl, 2)
        grid.Children.Add(options_lbl)

        self.cb_shared_coords = CheckBox(Content="Use Shared Coordinates")
        self.cb_shared_coords.IsChecked = True
        self.cb_shared_coords.Margin = Thickness(0, 2, 0, 2)
        Grid.SetRow(self.cb_shared_coords, 4)
        Grid.SetColumnSpan(self.cb_shared_coords, 2)
        grid.Children.Add(self.cb_shared_coords)

        self.cb_merged_views = CheckBox(Content="Merge Views")
        self.cb_merged_views.IsChecked = True
        self.cb_merged_views.Margin = Thickness(0, 2, 0, 2)
        Grid.SetRow(self.cb_merged_views, 5)
        Grid.SetColumnSpan(self.cb_merged_views, 2)
        grid.Children.Add(self.cb_merged_views)

        self.cb_hide_unref_tags = CheckBox(Content="Hide Unreferenced View Tags")
        self.cb_hide_unref_tags.IsChecked = True
        self.cb_hide_unref_tags.Margin = Thickness(0, 2, 0, 2)
        Grid.SetRow(self.cb_hide_unref_tags, 6)
        Grid.SetColumnSpan(self.cb_hide_unref_tags, 2)
        grid.Children.Add(self.cb_hide_unref_tags)

        self.cb_export_areas = CheckBox(Content="Export Room, space and area boundaries")
        self.cb_export_areas.IsChecked = False
        self.cb_export_areas.Margin = Thickness(0, 2, 0, 2)
        Grid.SetRow(self.cb_export_areas, 7)
        Grid.SetColumnSpan(self.cb_export_areas, 2)
        grid.Children.Add(self.cb_export_areas)

        button_panel = StackPanel(Orientation=Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center)
        button_panel.Margin = Thickness(0, 12, 0, 0)

        ok_btn = Button(Content="Export", Width=90, Height=28, Margin=Thickness(8, 0, 8, 0))
        ok_btn.Click += self.ok_clicked
        button_panel.Children.Add(ok_btn)

        cancel_btn = Button(Content="Cancel", Width=90, Height=28, Margin=Thickness(8, 0, 8, 0))
        cancel_btn.Click += self.cancel_clicked
        button_panel.Children.Add(cancel_btn)

        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        Grid.SetRow(button_panel, 8)
        Grid.SetColumnSpan(button_panel, 2)
        grid.Children.Add(button_panel)

        self.Content = grid

    def browse_clicked(self, sender, args):
        dlg = FolderBrowserDialog()
        dlg.Description = "Select output folder for DWG export"
        if os.path.isdir(self.folder_box.Text):
            dlg.SelectedPath = self.folder_box.Text

        if dlg.ShowDialog() == DialogResult.OK:
            self.folder_box.Text = dlg.SelectedPath

    def ok_clicked(self, sender, args):
        folder = self.folder_box.Text.strip()

        if not folder:
            TaskDialog.Show("Export DWG", "Please select an output folder.")
            return

        if not os.path.isdir(folder):
            TaskDialog.Show("Export DWG", "Selected folder does not exist.")
            return

        self.selected_folder = folder
        self.export_settings = {
            "SharedCoords": bool(self.cb_shared_coords.IsChecked),
            "MergedViews": bool(self.cb_merged_views.IsChecked),
            "HideUnreferenceViewTags": bool(self.cb_hide_unref_tags.IsChecked),
            "ExportingAreas": bool(self.cb_export_areas.IsChecked)
        }

        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


def get_selected_views_from_user_selection():
    selected_ids = list(uidoc.Selection.GetElementIds())
    selected_views = []

    if selected_ids:
        for eid in selected_ids:
            element = doc.GetElement(eid)
            if is_valid_export_view(element):
                selected_views.append(element)

    return selected_views


def prompt_user_to_select_views():
    all_views = [
        v for v in FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType()
        if is_valid_export_view(v)
    ]

    if not all_views:
        TaskDialog.Show("Export DWG", "No valid floor plan or ceiling plan views found in this document.")
        return []

    form = ViewSelectionDialog(
        all_views,
        "Select Views to Export",
        "Search and select floor plan / ceiling plan views:"
    )

    result = form.ShowDialog()
    if result:
        return form.selected_views

    return []


def prompt_for_export_folder_and_settings():
    dlg = ExportFolderAndOptionsDialog()
    if dlg.ShowDialog():
        return dlg.selected_folder, dlg.export_settings
    return None, None


def get_dwg_options(settings):
    dwg_options = Autodesk.Revit.DB.DWGExportOptions()

    dwg_options.SharedCoords = settings.get("SharedCoords", True)
    dwg_options.MergedViews = settings.get("MergedViews", True)
    dwg_options.HideUnreferenceViewTags = settings.get("HideUnreferenceViewTags", True)
    dwg_options.ExportingAreas = settings.get("ExportingAreas", False)

    # Fixed settings from your original script
    dwg_options.LineScaling = Autodesk.Revit.DB.LineScaling.PaperSpace
    dwg_options.TargetUnit = Autodesk.Revit.DB.ExportUnit.Inch

    return dwg_options


def export_view_to_dwg(view, output_folder, export_settings):
    try:
        dwg_options = get_dwg_options(export_settings)

        view_set = System.Collections.Generic.List[Autodesk.Revit.DB.ElementId]()
        view_set.Add(view.Id)

        file_name = sanitize_filename(view.Name)
        doc.Export(output_folder, file_name, view_set, dwg_options)

        return True, file_name + ".dwg"
    except Exception as ex:
        return False, "{} : {}".format(view.Name, str(ex))


def main():
    selected_views = get_selected_views_from_user_selection()

    if not selected_views:
        selected_views = prompt_user_to_select_views()

    if not selected_views:
        TaskDialog.Show("Export DWG", "No views selected. Operation cancelled.")
        return

    output_folder, export_settings = prompt_for_export_folder_and_settings()
    if not output_folder:
        TaskDialog.Show("Export DWG", "No output folder selected. Operation cancelled.")
        return

    exported = []
    failed = []

    for view in selected_views:
        success, result = export_view_to_dwg(view, output_folder, export_settings)
        if success:
            exported.append(result)
        else:
            failed.append(result)

    msg = []
    msg.append("DWG export complete.")
    msg.append("")
    msg.append("Views exported: {}".format(len(exported)))
    msg.append("Views failed: {}".format(len(failed)))

    if exported:
        msg.append("")
        msg.append("Exported:")
        msg.extend(exported[:15])

    if failed:
        msg.append("")
        msg.append("Failed:")
        msg.extend(failed[:15])

    TaskDialog.Show("Export DWG", "\n".join(msg))


if __name__ == "__main__":
    main()