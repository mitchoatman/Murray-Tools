# -*- coding: utf-8 -*-
import os
import clr
import System

from pyrevit import revit, DB, forms
from pyrevit.revit.db import query
from pyrevit.compat import get_elementid_value_func
import xlsxwriter

# WPF
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from System.Windows.Media import FontFamily

from System.Windows import (
    Window, Thickness, ResizeMode, HorizontalAlignment,
    GridLength, GridUnitType, SizeToContent
)
from System.Windows.Controls import (
    Label as WpfLabel,
    TextBox as WpfTextBox,
    Button as WpfButton,
    ScrollViewer,
    StackPanel,
    Grid,
    Orientation,
    CheckBox,
    TextBlock,
    ListBox,
    SelectionMode
)

doc = revit.doc
get_eid_value = get_elementid_value_func()


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def get_param_readonly(tb, param_name):
    p = get_param_from_tb_or_symbol(tb, param_name)
    if not p:
        return True
    try:
        return p.IsReadOnly
    except Exception:
        return True


def family_param_is_readonly(title_blocks, param_name):
    found_any = False
    for tb in title_blocks:
        p = get_param_from_tb_or_symbol(tb, param_name)
        if p:
            found_any = True
            try:
                if not p.IsReadOnly:
                    return False
            except Exception:
                pass
    return True if found_any else True

def show_message(msg, title="Export Title Blocks"):
    TaskDialog.Show(title, msg)


def ask_yes_no(msg, title="Export Title Blocks"):
    td = TaskDialog(title)
    td.MainInstruction = msg
    td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    td.DefaultButton = TaskDialogResult.Yes
    return td.Show() == TaskDialogResult.Yes

def sanitize_filename(name):
    invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    safe = name or "Export"
    for ch in invalid:
        safe = safe.replace(ch, '_')
    safe = safe.strip().rstrip('.')
    return safe or "Export"


def get_project_export_name():
    try:
        project_name = query.get_document_clean_name(doc)
    except Exception:
        project_name = ""

    if not project_name or project_name == "File Not Saved":
        try:
            project_name = os.path.splitext(doc.Title)[0]
        except Exception:
            project_name = "Project"

    return sanitize_filename(project_name)


def build_default_export_filename(selected_family_names):
    project_name = get_project_export_name()

    if len(selected_family_names) == 1:
        tb_name = selected_family_names[0]
    else:
        tb_name = "MULTI_TBLOCKS"

    tb_name = sanitize_filename(tb_name)
    return "{}_{}.xlsx".format(project_name, tb_name)

def get_tb_family_name(tb):
    try:
        if tb.Symbol and tb.Symbol.Family:
            return tb.Symbol.Family.Name
    except Exception:
        pass
    return "Unnamed"


def sanitize_sheet_name(name):
    invalid = ['\\', '/', ':', '*', '?', '[', ']']
    safe = name or "Sheet"
    for ch in invalid:
        safe = safe.replace(ch, '_')
    safe = safe.strip()
    if not safe:
        safe = "Sheet"
    return safe[:31]


def make_unique_sheet_name(name, used_names):
    base = sanitize_sheet_name(name)
    if base not in used_names:
        used_names.add(base)
        return base

    idx = 2
    while True:
        suffix = "_{0}".format(idx)
        candidate = base[:31 - len(suffix)] + suffix
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        idx += 1


def get_param_from_tb_or_symbol(tb, param_name):
    try:
        p = tb.LookupParameter(param_name)
        if p:
            return p
    except Exception:
        pass

    try:
        if tb.Symbol:
            p = tb.Symbol.LookupParameter(param_name)
            if p:
                return p
    except Exception:
        pass

    return None


def get_param_value(tb, param_name):
    p = get_param_from_tb_or_symbol(tb, param_name)
    if not p:
        return ""

    try:
        st = p.StorageType

        if st == DB.StorageType.String:
            return p.AsString() or ""

        elif st == DB.StorageType.Double:
            try:
                display = p.AsValueString()
                if display:
                    return display
            except Exception:
                pass
            try:
                return p.AsDouble()
            except Exception:
                return ""

        elif st == DB.StorageType.Integer:
            try:
                return p.AsInteger()
            except Exception:
                return ""

        elif st == DB.StorageType.ElementId:
            try:
                eid = p.AsElementId()
                if eid:
                    return get_eid_value(eid)
            except Exception:
                return ""
            return ""

    except Exception:
        return ""

    return ""


def get_sort_value(tb, param_name):
    value = get_param_value(tb, param_name)
    try:
        if value is None:
            return ("", get_eid_value(tb.Id))
        return (str(value).lower(), get_eid_value(tb.Id))
    except Exception:
        return ("", get_eid_value(tb.Id))


def get_all_parameters(title_blocks):
    param_names = set()
    processed_family_ids = set()

    for tb in title_blocks:
        try:
            for param in tb.Parameters:
                try:
                    pname = param.Definition.Name
                    if pname:
                        param_names.add(pname)
                except Exception:
                    pass
        except Exception:
            pass

        try:
            if tb.Symbol:
                for param in tb.Symbol.Parameters:
                    try:
                        pname = param.Definition.Name
                        if pname:
                            param_names.add(pname)
                    except Exception:
                        pass
        except Exception:
            pass

        try:
            fam = tb.Symbol.Family if tb.Symbol and tb.Symbol.Family else None
            if fam:
                fam_id = get_eid_value(fam.Id)
                if fam_id not in processed_family_ids:
                    family_doc = None
                    try:
                        family_doc = doc.EditFamily(fam)
                        if family_doc:
                            for fparam in family_doc.FamilyManager.Parameters:
                                try:
                                    pname = fparam.Definition.Name
                                    if pname:
                                        param_names.add(pname)
                                except Exception:
                                    pass
                    except Exception:
                        pass
                    finally:
                        try:
                            if family_doc:
                                family_doc.Close(False)
                        except Exception:
                            pass

                    processed_family_ids.add(fam_id)
        except Exception:
            pass

    return sorted(param_names)


def export_titleblocks_to_excel(title_blocks, filepath, selected_params):
    tb_by_family = {}
    used_sheet_names = set()

    for tb in title_blocks:
        family_name = get_tb_family_name(tb)
        if family_name not in tb_by_family:
            tb_by_family[family_name] = []
        tb_by_family[family_name].append(tb)

    workbook = xlsxwriter.Workbook(filepath)

    # Formats
    header_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#D9E1F2",
        "border": 1,
        "locked": True
    })

    header_editable_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#C6E0B4",
        "border": 1,
        "locked": True
    })

    header_readonly_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#FF3131",
        "font_color": "#FFFFFF",
        "border": 1,
        "locked": True
    })

    elementid_header_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#FFBD80",
        "border": 1,
        "locked": True
    })

    locked_fmt = workbook.add_format({
        "locked": True,
        "bg_color": "#F2F2F2"
    })

    unlocked_fmt = workbook.add_format({
        "locked": False,
        "bg_color": "#FFF2CC"
    })

    missing_fmt = workbook.add_format({
        "locked": True,
        "bg_color": "#E7E6E6",
        "font_color": "#7F7F7F"
    })

    for family_name in sorted(tb_by_family.keys()):
        tbs = tb_by_family[family_name]
        sheet_name = make_unique_sheet_name(family_name, used_sheet_names)
        ws = workbook.add_worksheet(sheet_name)

        if selected_params:
            first_param = selected_params[0]
            sorted_tbs = sorted(tbs, key=lambda x: get_sort_value(x, first_param))
        else:
            sorted_tbs = sorted(tbs, key=lambda x: get_eid_value(x.Id))

        # Header row
        ws.write(0, 0, "ElementId", elementid_header_fmt)

        for col_idx, param_name in enumerate(selected_params, start=1):
            if family_param_is_readonly(tbs, param_name):
                hdr_fmt = header_readonly_fmt
            else:
                hdr_fmt = header_editable_fmt
            ws.write(0, col_idx, param_name, hdr_fmt)

        # Data rows
        max_widths = [len("ElementId")] + [len(p) for p in selected_params]

        for row_idx, tb in enumerate(sorted_tbs, start=1):
            eid_text = str(get_eid_value(tb.Id))
            ws.write(row_idx, 0, eid_text, locked_fmt)
            if len(eid_text) > max_widths[0]:
                max_widths[0] = min(len(eid_text), 50)

            for col_idx, param_name in enumerate(selected_params, start=1):
                p = get_param_from_tb_or_symbol(tb, param_name)

                if not p:
                    value = ""
                    fmt = missing_fmt
                else:
                    value = get_param_value(tb, param_name)
                    try:
                        fmt = locked_fmt if p.IsReadOnly else unlocked_fmt
                    except Exception:
                        fmt = locked_fmt

                if value is None:
                    value = ""

                ws.write(row_idx, col_idx, value, fmt)

                val_len = len(str(value))
                if val_len > max_widths[col_idx]:
                    max_widths[col_idx] = min(val_len, 50)

        # Column widths
        for col_idx, width in enumerate(max_widths):
            ws.set_column(col_idx, col_idx, width + 3)

        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(sorted_tbs), len(selected_params))

        ws.protect(
            "",
            {
                "autofilter": True,
                "sort": True,
                "format_cells": True,
                "format_columns": True,
                "format_rows": True,
                "select_locked_cells": True,
                "select_unlocked_cells": True,
            },
        )

    workbook.close()

# -----------------------------------------------------------------------------
# WPF family selection dialog
# -----------------------------------------------------------------------------

class TitleBlockFamilySelectionWindow(Window):
    def __init__(self, family_names):
        self.selected_family_names = []
        self.family_list = sorted(family_names)
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Select Title Block Families"
        self.Width = 400
        self.Height = 400
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(5)

        for i in range(4):
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        self.label = WpfLabel(Content="Select title block families to export:")
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        self.search_box = WpfTextBox(Height=22, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        self.checkbox_panel = StackPanel(Orientation=Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel)
        scroll_viewer.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll_viewer.Margin = Thickness(0, 1, 0, 1)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.family_list)

        button_panel = StackPanel(
            Orientation=Orientation.Horizontal,
            HorizontalAlignment=HorizontalAlignment.Center,
            Margin=Thickness(0, 10, 0, 10)
        )

        self.select_button = WpfButton(
            Content="Select",
            FontFamily=FontFamily("Arial"),
            FontSize=12,
            Height=25,
            Margin=Thickness(10, 0, 10, 0),
            Width=60
        )
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = WpfButton(
            Content="Check All",
            FontFamily=FontFamily("Arial"),
            FontSize=12,
            Height=25,
            Margin=Thickness(10, 0, 10, 0),
            Width=80
        )
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        self.Content = grid

    def update_checkboxes(self, family_names):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []

        for family_name in family_names:
            tb_display = TextBlock()
            tb_display.Text = family_name

            checkbox = CheckBox(Content=tb_display)
            checkbox.Tag = family_name
            checkbox.Click += self.checkbox_clicked

            if family_name in self.selected_family_names:
                checkbox.IsChecked = True

            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [x for x in self.family_list if search_text in x.lower()]
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self._refresh_selected()

    def checkbox_clicked(self, sender, args):
        self._refresh_selected()

    def _refresh_selected(self):
        current_visible = [cb.Tag for cb in self.checkboxes if cb.IsChecked]
        preserved_hidden = [x for x in self.selected_family_names if x not in [cb.Tag for cb in self.checkboxes]]
        self.selected_family_names = preserved_hidden + current_visible

    def select_clicked(self, sender, args):
        self._refresh_selected()
        self.DialogResult = True
        self.Close()


# -----------------------------------------------------------------------------
# Parameter ordering dialog
# -----------------------------------------------------------------------------

class ParameterSelectionDialog(Window):
    def __init__(self, param_names):
        self.selected_params = []
        self.all_params = sorted(param_names)
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Select and Order Parameters"
        self.Width = 720
        self.Height = 500
        self.MinWidth = 720
        self.MinHeight = 500
        self.ResizeMode = ResizeMode.CanResize
        self.SizeToContent = SizeToContent.Manual
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        main_grid = Grid()
        main_grid.Margin = Thickness(10)

        # -------------------------
        # Rows
        # -------------------------
        row0 = System.Windows.Controls.RowDefinition()
        row0.Height = GridLength.Auto
        main_grid.RowDefinitions.Add(row0)

        row1 = System.Windows.Controls.RowDefinition()
        row1.Height = GridLength(1, GridUnitType.Star)
        main_grid.RowDefinitions.Add(row1)

        row2 = System.Windows.Controls.RowDefinition()
        row2.Height = GridLength.Auto
        main_grid.RowDefinitions.Add(row2)

        # -------------------------
        # Columns
        # left list | button stack | right list
        # -------------------------
        col0 = System.Windows.Controls.ColumnDefinition()
        col0.Width = GridLength(1, GridUnitType.Star)
        main_grid.ColumnDefinitions.Add(col0)

        col1 = System.Windows.Controls.ColumnDefinition()
        col1.Width = GridLength.Auto
        main_grid.ColumnDefinitions.Add(col1)

        col2 = System.Windows.Controls.ColumnDefinition()
        col2.Width = GridLength(1, GridUnitType.Star)
        main_grid.ColumnDefinitions.Add(col2)

        # -------------------------
        # Header labels
        # -------------------------
        self.label_available = WpfLabel(
            Content="Available Parameters",
            FontFamily=FontFamily("Arial"),
            FontSize=14,
            Margin=Thickness(0, 0, 10, 5),
            HorizontalAlignment=HorizontalAlignment.Center,
            HorizontalContentAlignment=HorizontalAlignment.Center
        )
        Grid.SetRow(self.label_available, 0)
        Grid.SetColumn(self.label_available, 0)
        main_grid.Children.Add(self.label_available)

        self.label_selected = WpfLabel(
            Content="Selected Parameters (Order Matters)",
            FontFamily=FontFamily("Arial"),
            FontSize=14,
            Margin=Thickness(10, 0, 0, 5),
            HorizontalAlignment=HorizontalAlignment.Stretch,
            HorizontalContentAlignment=HorizontalAlignment.Center
        )
        Grid.SetRow(self.label_selected, 0)
        Grid.SetColumn(self.label_selected, 2)
        main_grid.Children.Add(self.label_selected)

        # -------------------------
        # Left list
        # -------------------------
        self.available_listbox = ListBox()
        self.available_listbox.SelectionMode = SelectionMode.Extended
        self.available_listbox.FontFamily = FontFamily("Arial")
        self.available_listbox.FontSize = 12
        self.available_listbox.Margin = Thickness(0, 0, 10, 0)

        for p in self.all_params:
            self.available_listbox.Items.Add(p)

        Grid.SetRow(self.available_listbox, 1)
        Grid.SetColumn(self.available_listbox, 0)
        main_grid.Children.Add(self.available_listbox)

        # -------------------------
        # Middle buttons
        # -------------------------
        middle_panel = StackPanel()
        middle_panel.Orientation = Orientation.Vertical
        middle_panel.HorizontalAlignment = HorizontalAlignment.Center
        middle_panel.VerticalAlignment = System.Windows.VerticalAlignment.Center
        middle_panel.Margin = Thickness(10, 0, 10, 0)

        self.add_button = WpfButton(Content=">>", Width=70, Height=30, Margin=Thickness(0, 0, 0, 10))
        self.add_button.Click += self.add_selected
        middle_panel.Children.Add(self.add_button)

        self.remove_button = WpfButton(Content="<<", Width=70, Height=30, Margin=Thickness(0, 0, 0, 20))
        self.remove_button.Click += self.remove_selected
        middle_panel.Children.Add(self.remove_button)

        self.up_button = WpfButton(Content="Up", Width=70, Height=30, Margin=Thickness(0, 0, 0, 10))
        self.up_button.Click += self.move_up
        middle_panel.Children.Add(self.up_button)

        self.down_button = WpfButton(Content="Down", Width=70, Height=30)
        self.down_button.Click += self.move_down
        middle_panel.Children.Add(self.down_button)

        Grid.SetRow(middle_panel, 1)
        Grid.SetColumn(middle_panel, 1)
        main_grid.Children.Add(middle_panel)

        # -------------------------
        # Right list
        # -------------------------
        self.selected_listbox = ListBox()
        self.selected_listbox.SelectionMode = SelectionMode.Extended
        self.selected_listbox.FontFamily = FontFamily("Arial")
        self.selected_listbox.FontSize = 12
        self.selected_listbox.Margin = Thickness(10, 0, 0, 0)

        Grid.SetRow(self.selected_listbox, 1)
        Grid.SetColumn(self.selected_listbox, 2)
        main_grid.Children.Add(self.selected_listbox)

        # -------------------------
        # Bottom buttons
        # -------------------------
        bottom_panel = StackPanel()
        bottom_panel.Orientation = Orientation.Horizontal
        bottom_panel.HorizontalAlignment = HorizontalAlignment.Center
        bottom_panel.Margin = Thickness(0, 12, 0, 0)

        self.ok_button = WpfButton(Content="OK", Width=90, Height=30, Margin=Thickness(8, 0, 8, 0))
        self.ok_button.Click += self.ok_clicked
        bottom_panel.Children.Add(self.ok_button)

        self.cancel_button = WpfButton(Content="Cancel", Width=90, Height=30, Margin=Thickness(8, 0, 8, 0))
        self.cancel_button.Click += self.cancel_clicked
        bottom_panel.Children.Add(self.cancel_button)

        Grid.SetRow(bottom_panel, 2)
        Grid.SetColumnSpan(bottom_panel, 3)
        main_grid.Children.Add(bottom_panel)

        self.Content = main_grid

    def add_selected(self, sender, args):
        selected_items = list(self.available_listbox.SelectedItems)
        for item in selected_items:
            if not self.selected_listbox.Items.Contains(item):
                self.selected_listbox.Items.Add(item)

        for item in selected_items:
            self.available_listbox.Items.Remove(item)

    def remove_selected(self, sender, args):
        selected_items = list(self.selected_listbox.SelectedItems)
        for item in selected_items:
            self.selected_listbox.Items.Remove(item)
            if not self.available_listbox.Items.Contains(item):
                self.available_listbox.Items.Add(item)

        available_sorted = sorted([self.available_listbox.Items[i] for i in range(self.available_listbox.Items.Count)])
        self.available_listbox.Items.Clear()
        for item in available_sorted:
            self.available_listbox.Items.Add(item)

    def move_up(self, sender, args):
        selected_items = list(self.selected_listbox.SelectedItems)
        if not selected_items:
            return

        indices = sorted([self.selected_listbox.Items.IndexOf(item) for item in selected_items])

        for index in indices:
            if index > 0:
                item = self.selected_listbox.Items[index]
                above_item = self.selected_listbox.Items[index - 1]
                self.selected_listbox.Items[index - 1] = item
                self.selected_listbox.Items[index] = above_item

        self.selected_listbox.SelectedItems.Clear()
        for item in selected_items:
            self.selected_listbox.SelectedItems.Add(item)

    def move_down(self, sender, args):
        selected_items = list(self.selected_listbox.SelectedItems)
        if not selected_items:
            return

        indices = sorted(
            [self.selected_listbox.Items.IndexOf(item) for item in selected_items],
            reverse=True
        )

        for index in indices:
            if index < self.selected_listbox.Items.Count - 1:
                item = self.selected_listbox.Items[index]
                below_item = self.selected_listbox.Items[index + 1]
                self.selected_listbox.Items[index + 1] = item
                self.selected_listbox.Items[index] = below_item

        self.selected_listbox.SelectedItems.Clear()
        for item in selected_items:
            self.selected_listbox.SelectedItems.Add(item)

    def ok_clicked(self, sender, args):
        self.selected_params = [self.selected_listbox.Items[i] for i in range(self.selected_listbox.Items.Count)]
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

def main():
    title_blocks = list(
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    if not title_blocks:
        forms.alert("No title blocks found in the project.", warn_icon=True)
        return

    tb_dict = {}
    for tb in title_blocks:
        family_name = get_tb_family_name(tb)
        tb_dict.setdefault(family_name, []).append(tb)

    tb_form = TitleBlockFamilySelectionWindow(tb_dict.keys())
    if not tb_form.ShowDialog() or not tb_form.selected_family_names:
        show_message("No title block families selected.")
        return

    selected_title_blocks = []
    for family_name in tb_form.selected_family_names:
        selected_title_blocks.extend(tb_dict[family_name])

    param_names = get_all_parameters(selected_title_blocks)
    if not param_names:
        show_message("No parameters found.")
        return

    param_dialog = ParameterSelectionDialog(param_names)
    result = param_dialog.ShowDialog()

    if not result or not param_dialog.selected_params:
        show_message("No parameters selected.")
        return

    default_filename = build_default_export_filename(tb_form.selected_family_names)

    filepath = forms.save_file(
        file_ext="xlsx",
        default_name=default_filename,
        title="Save Title Block Export"
    )

    if not filepath:
        return

    if not filepath.lower().endswith(".xlsx"):
        filepath += ".xlsx"

    try:
        if os.path.exists(filepath):
            os.remove(filepath)
    except Exception as ex:
        show_message("Export failed:\n{0}".format(str(ex)))
        return

    try:
        export_titleblocks_to_excel(
            selected_title_blocks,
            filepath,
            param_dialog.selected_params
        )
    except Exception as ex:
        forms.alert("Export failed:\n{0}".format(str(ex)), warn_icon=True)
        return

    open_now = ask_yes_no("Export complete.\n\nOpen the file now?")

    if open_now:
        try:
            os.startfile(filepath)
        except Exception:
            show_message("File was created but could not be opened automatically.")


if __name__ == "__main__":
    main()