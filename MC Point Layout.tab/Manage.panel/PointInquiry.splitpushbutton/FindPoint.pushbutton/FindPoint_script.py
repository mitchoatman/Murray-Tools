# -*- coding: utf-8 -*-
import os
import clr
import System

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference("System.Core")

from System import Action
import System.Windows.Threading

from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, FontWeights, GridLength, GridUnitType, Setter, Style
from System.Windows.Controls import Grid, RowDefinition, ColumnDefinition, Label, TextBox, Button, DataGrid, DataGridTextColumn, DataGridSelectionMode, Orientation, StackPanel, DataGridLength, Control
from System.Windows.Controls.Primitives import DataGridColumnHeader, DataGridRowHeader
from System.Windows.Data import Binding, CollectionViewSource
from System.Windows.Interop import WindowInteropHelper
from System.Collections.Generic import List
from System.Windows.Media import SolidColorBrush, Color as MediaColor

from Autodesk.Revit.DB import FilteredElementCollector, ElementId, Transaction
from Autodesk.Revit.UI import UIApplication, TaskDialog, UIThemeManager, UITheme

from Parameters.Add_SharedParameters import Shared_Params


# -----------------------------------------------------------------------------
# Setup
# -----------------------------------------------------------------------------
Shared_Params()

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app = UIApplication(doc.Application)

FOLDER_NAME = r"C:\Temp"
FILE_PATH = os.path.join(FOLDER_NAME, "Ribbon_TSPointNumber.txt")
PARAM_NAME = "TS_Point_Number"
DESC_PARAM_NAME = "TS_Point_Description"


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def natural_key(s):
    import re
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'([0-9]+)', s)]


def ensure_input_file():
    if not os.path.exists(FOLDER_NAME):
        os.makedirs(FOLDER_NAME)

    if not os.path.exists(FILE_PATH):
        with open(FILE_PATH, 'w') as f:
            f.write('')


def read_previous_input():
    ensure_input_file()
    with open(FILE_PATH, 'r') as f:
        return f.read().strip()


def write_previous_input(value):
    ensure_input_file()
    with open(FILE_PATH, 'w') as f:
        f.write(value or '')


def get_revit_window_handle():
    try:
        return uidoc.Application.MainWindowHandle
    except:
        try:
            return UIApplication(doc.Application).MainWindowHandle
        except:
            return System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle


def get_point_data_in_active_view():
    data_list = []
    collector = FilteredElementCollector(doc, doc.ActiveView.Id)

    for elem in collector:
        param = elem.LookupParameter(PARAM_NAME)
        if param and param.HasValue:
            val = param.AsString()
            if val:
                desc_param = elem.LookupParameter(DESC_PARAM_NAME)
                desc = ""
                if desc_param and desc_param.HasValue:
                    desc = desc_param.AsString() or ""
                
                data_list.append({
                    "POINT NUMBER": val,
                    "DESCRIPTION": desc,
                    "_Elem": elem
                })

    # Sort by point number naturally
    data_list = sorted(data_list, key=lambda x: natural_key(x["POINT NUMBER"]))
    return data_list


def show_elements(elements):
    if not elements:
        return

    ids = List[ElementId]()
    for elem in elements:
        ids.Add(elem.Id)

    uidoc.Selection.SetElementIds(ids)
    uidoc.ShowElements(ids)


# -----------------------------------------------------------------------------
# Window
# -----------------------------------------------------------------------------
class PointDataWindow(Window):
    def __init__(self, default_value, point_data, revit_window_handle):
        Window.__init__(self)

        self.Title = "TS Point Number Locator"
        self.Width = 360
        self.Height = 480
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        self.point_data = point_data
        self.is_initialized = False
        self.is_filtering = False

        self.apply_revit_theme()
        self.initialize_components(default_value)

        try:
            WindowInteropHelper(self).Owner = revit_window_handle
        except:
            pass

        self.is_initialized = True

    def apply_revit_theme(self):
        try:
            current_theme = UIThemeManager.CurrentTheme
            is_dark = (current_theme == UITheme.Dark)

            bg_hex = "#FF3B4453" if is_dark else "#FFF5F5F5"
            text_hex = "#FFDFDFDF" if is_dark else "#FF333333"
            subtext_hex = "#FF999999" if is_dark else "#FF666666"
            ctrl_bg_hex = "#FF222933" if is_dark else "#FFFFFFFF"
            btn_bg_hex = "#FF222933" if is_dark else "#FFEFEFEF"
            border_hex = "#FF363B40" if is_dark else "#FFD0D0D0"
            grid_cell_bg_hex = "#FF2A323D" if is_dark else "#FFFFFFFF"
            grid_header_bg_hex = "#FF323B48" if is_dark else "#FFEFEFEF"

            res = self.Resources
            res["PageBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["TextForegroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(text_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["SubTextForegroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(subtext_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["ControlBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(ctrl_bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["ButtonBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(btn_bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["BorderColorBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(border_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["GridCellBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(grid_cell_bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["GridHeaderBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(grid_header_bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))

            self.Background = res["PageBackgroundBrush"]
        except Exception:
            pass

    def initialize_components(self, default_value):
        grid = Grid()
        grid.Margin = Thickness(10)
        self.Content = grid

        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1.0, GridUnitType.Star)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

        grid.ColumnDefinitions.Add(
            ColumnDefinition(Width=GridLength(1.0, GridUnitType.Star))
        )

        row_index = 0

        self.label = Label()
        self.label.Content = "Find Point Data:"
        self.label.FontWeight = FontWeights.Bold
        self.label.Margin = Thickness(0, 0, 0, 3)
        self.label.SetResourceReference(Label.ForegroundProperty, "TextForegroundBrush")
        Grid.SetRow(self.label, row_index)
        grid.Children.Add(self.label)
        row_index += 1

        self.textbox = TextBox()
        self.textbox.Text = default_value or ""
        self.textbox.Height = 26
        self.textbox.VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        self.textbox.Margin = Thickness(0, 0, 0, 8)
        self.textbox.SetResourceReference(TextBox.BackgroundProperty, "ControlBackgroundBrush")
        self.textbox.SetResourceReference(TextBox.ForegroundProperty, "TextForegroundBrush")
        self.textbox.SetResourceReference(TextBox.BorderBrushProperty, "BorderColorBrush")
        self.textbox.TextChanged += self.on_text_changed
        Grid.SetRow(self.textbox, row_index)
        grid.Children.Add(self.textbox)
        row_index += 1

        # Collection View Source for DataGrid Filtering
        self.cvs = CollectionViewSource()
        self.cvs.Source = self.point_data
        self.cvs.Filter += self.filter_rows

        self.datagrid = DataGrid()
        self.datagrid.AutoGenerateColumns = False
        self.datagrid.IsReadOnly = False
        self.datagrid.SelectionMode = DataGridSelectionMode.Single
        self.datagrid.SetResourceReference(DataGrid.BackgroundProperty, "GridCellBackgroundBrush")
        self.datagrid.SetResourceReference(DataGrid.ForegroundProperty, "TextForegroundBrush")
        self.datagrid.SetResourceReference(DataGrid.BorderBrushProperty, "BorderColorBrush")
        self.datagrid.SetResourceReference(DataGrid.RowBackgroundProperty, "GridCellBackgroundBrush")
        self.datagrid.SetResourceReference(DataGrid.AlternatingRowBackgroundProperty, "GridCellBackgroundBrush")
        self.datagrid.ColumnHeaderStyle = self.create_header_style()
        self.datagrid.RowHeaderStyle = self.create_row_header_style()
        self.datagrid.SelectionChanged += self.on_selection_changed
        self.datagrid.CellEditEnding += self.on_cell_edit_ending
        self.datagrid.Margin = Thickness(0, 0, 0, 10)

        # Columns with explicit DataGridLength widths (~1.5x / ~2.5x proportional ratio)
        col_num = DataGridTextColumn()
        col_num.Header = "Point Number"
        col_num.Binding = Binding("[POINT NUMBER]")
        col_num.IsReadOnly = True
        col_num.Width = DataGridLength(100.0)
        self.datagrid.Columns.Add(col_num)

        col_desc = DataGridTextColumn()
        col_desc.Header = "Description"
        col_desc.Binding = Binding("[DESCRIPTION]")
        col_desc.IsReadOnly = False
        col_desc.Width = DataGridLength(200.0)
        self.datagrid.Columns.Add(col_desc)

        self.datagrid.ItemsSource = self.cvs.View
        Grid.SetRow(self.datagrid, row_index)
        grid.Children.Add(self.datagrid)
        row_index += 1

        # Close Button Panel
        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 5, 0, 0)
        Grid.SetRow(button_panel, row_index)
        grid.Children.Add(button_panel)

        self.close_button = Button()
        self.close_button.Content = "Close"
        self.close_button.Width = 85
        self.close_button.Height = 28
        self.close_button.SetResourceReference(Button.BackgroundProperty, "ButtonBackgroundBrush")
        self.close_button.SetResourceReference(Button.ForegroundProperty, "TextForegroundBrush")
        self.close_button.SetResourceReference(Button.BorderBrushProperty, "BorderColorBrush")
        self.close_button.Click += self.on_close_click
        button_panel.Children.Add(self.close_button)

        self.KeyDown += self.on_key_down

    def create_header_style(self):
        style = Style(DataGridColumnHeader)
        style.Setters.Add(Setter(Control.BackgroundProperty, self.Resources["GridHeaderBackgroundBrush"]))
        style.Setters.Add(Setter(Control.ForegroundProperty, self.Resources["TextForegroundBrush"]))
        style.Setters.Add(Setter(Control.BorderBrushProperty, self.Resources["BorderColorBrush"]))
        style.Setters.Add(Setter(Control.BorderThicknessProperty, Thickness(0, 0, 1, 1)))
        return style

    def create_row_header_style(self):
        style = Style(DataGridRowHeader)
        style.Setters.Add(Setter(Control.BackgroundProperty, self.Resources["GridHeaderBackgroundBrush"]))
        style.Setters.Add(Setter(Control.BorderBrushProperty, self.Resources["BorderColorBrush"]))
        style.Setters.Add(Setter(Control.BorderThicknessProperty, Thickness(0, 0, 1, 1)))
        return style

    def save_search_text(self):
        try:
            write_previous_input(self.textbox.Text.strip())
        except:
            pass

    def filter_rows(self, sender, args):
        try:
            query = self.textbox.Text.strip().lower()
            if not query:
                args.Accepted = True
                return

            item = args.Item
            point_num = str(item.get("POINT NUMBER", "")).lower()
            desc = str(item.get("DESCRIPTION", "")).lower()

            if query in point_num or query in desc:
                args.Accepted = True
            else:
                args.Accepted = False
        except:
            args.Accepted = True

    def on_text_changed(self, sender, event):
        self.is_filtering = True
        try:
            if self.cvs and self.cvs.View:
                self.cvs.View.Refresh()
        except:
            pass
        self.is_filtering = False

    def on_selection_changed(self, sender, event):
        if self.is_filtering or not self.is_initialized:
            return
        try:
            selected_item = self.datagrid.SelectedItem
            if selected_item:
                elem = selected_item.get("_Elem")
                if elem:
                    show_elements([elem])
        except:
            pass

    def on_cell_edit_ending(self, sender, args):
        try:
            row_item = args.Row.Item
            if not row_item:
                return
            
            # Extract the newly typed text from the editing textbox
            tb = args.EditingElement
            if hasattr(tb, "Text"):
                new_val = tb.Text
                row_item["DESCRIPTION"] = new_val
                
                # Write back immediately to Revit element
                elem = row_item.get("_Elem")
                if elem:
                    t = Transaction(doc, "Update Point Description")
                    t.Start()
                    p = elem.LookupParameter(DESC_PARAM_NAME)
                    if p and not p.IsReadOnly:
                        p.Set(new_val)
                    t.Commit()
        except Exception as ex:
            pass

    def on_close_click(self, sender, event):
        self.save_search_text()
        self.Close()

    def on_key_down(self, sender, event):
        if event.Key == System.Windows.Input.Key.Escape:
            self.save_search_text()
            self.Close()


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
previous_input = read_previous_input()
point_data = get_point_data_in_active_view()
revit_window_handle = get_revit_window_handle()

form = PointDataWindow(previous_input, point_data, revit_window_handle)
form.Show()

# Clear initial selection so nothing is highlighted on open
form.datagrid.UnselectAll()

disp = System.Windows.Threading.Dispatcher.CurrentDispatcher
while form.IsVisible:
    disp.Invoke(System.Windows.Threading.DispatcherPriority.Background, Action(lambda: None))

form.save_search_text()