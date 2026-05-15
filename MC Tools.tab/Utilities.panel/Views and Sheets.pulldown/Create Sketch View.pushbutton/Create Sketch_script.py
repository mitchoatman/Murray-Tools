# -*- coding: UTF-8 -*-
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, BuiltInCategory, Transaction, ViewSheet, ViewDuplicateOption, Viewport, ElementId
from Autodesk.Revit.UI.Selection import PickBoxStyle
from Autodesk.Revit.UI import TaskDialog
import sys, os, clr

clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

import System
from System.Windows import Window, Thickness, ResizeMode, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType, WindowStartupLocation
from System.Windows.Controls import Label as WpfLabel, TextBox as WpfTextBox, Button as WpfButton, ScrollViewer, StackPanel, Grid, CheckBox, RadioButton, RowDefinition, ColumnDefinition, ScrollBarVisibility
from System.Windows.Controls.Primitives import UniformGrid
from System.Windows.Media import FontFamily, Brushes

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector

if str(curview.ViewType) == 'DrawingSheet':
    TaskDialog.Show("Error", "This script cannot be run from a sheet view. Please open a model or detail view and try again.")
    sys.exit()

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_CreateSketch.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as the_file:
        line1 = ('SK-001' + '\n')
        line2 = 'SKETCH-001'
        the_file.writelines([line1, line2])

with open(filepath, 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]


class SketchDataForm(Window):
    def __init__(self, default_number, default_name):
        self.snumber = None
        self.sname = None
        self.InitializeComponents(default_number, default_name)

    def InitializeComponents(self, default_number, default_name):
        self.Title = "Sketch Data"
        self.Width = 325
        self.Height = 220
        self.MinWidth = 325
        self.MinHeight = 220
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        grid = Grid()
        grid.Margin = Thickness(10)

        for i in range(5):
            row = RowDefinition()
            row.Height = GridLength.Auto
            grid.RowDefinitions.Add(row)

        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition())

        self.label1 = WpfLabel()
        self.label1.Content = "Sheet Number:"
        self.label1.FontFamily = FontFamily("Arial")
        self.label1.FontSize = 14
        self.label1.Margin = Thickness(0, 0, 0, 4)
        Grid.SetRow(self.label1, 0)
        grid.Children.Add(self.label1)

        self.textBox1 = WpfTextBox()
        self.textBox1.Text = default_number
        self.textBox1.Height = 24
        self.textBox1.FontFamily = FontFamily("Arial")
        self.textBox1.FontSize = 12
        self.textBox1.Margin = Thickness(0, 0, 0, 8)
        Grid.SetRow(self.textBox1, 1)
        grid.Children.Add(self.textBox1)

        self.label2 = WpfLabel()
        self.label2.Content = "Sheet / View Name:"
        self.label2.FontFamily = FontFamily("Arial")
        self.label2.FontSize = 14
        self.label2.Margin = Thickness(0, 0, 0, 4)
        Grid.SetRow(self.label2, 2)
        grid.Children.Add(self.label2)

        self.textBox2 = WpfTextBox()
        self.textBox2.Text = default_name
        self.textBox2.Height = 24
        self.textBox2.FontFamily = FontFamily("Arial")
        self.textBox2.FontSize = 12
        self.textBox2.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.textBox2, 3)
        grid.Children.Add(self.textBox2)

        button_panel = UniformGrid()
        button_panel.Columns = 2
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 5, 0, 0)

        self.ok_button = WpfButton()
        self.ok_button.Content = "Set Data"
        self.ok_button.FontFamily = FontFamily("Arial")
        self.ok_button.FontSize = 12
        self.ok_button.Height = 28
        self.ok_button.Margin = Thickness(5, 0, 5, 0)
        self.ok_button.Click += self.on_ok
        button_panel.Children.Add(self.ok_button)

        self.cancel_button = WpfButton()
        self.cancel_button.Content = "Cancel"
        self.cancel_button.FontFamily = FontFamily("Arial")
        self.cancel_button.FontSize = 12
        self.cancel_button.Height = 28
        self.cancel_button.Margin = Thickness(5, 0, 5, 0)
        self.cancel_button.Background = Brushes.LightGray
        self.cancel_button.Click += self.on_cancel
        button_panel.Children.Add(self.cancel_button)

        Grid.SetRow(button_panel, 4)
        grid.Children.Add(button_panel)

        self.Content = grid
        self.Loaded += self.on_loaded

    def on_loaded(self, sender, args):
        self.textBox1.Focus()
        self.textBox1.SelectAll()

    def on_ok(self, sender, args):
        self.snumber = self.textBox1.Text
        self.sname = self.textBox2.Text
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


class TitleblockSelectionFilter(Window):
    def __init__(self, titleblocks):
        self.selected_titleblock = None
        self.titleblock_list = sorted(
            titleblocks,
            key=lambda x: "{} - {}".format(
                x.FamilyName,
                x.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            )
        )
        self.radio_buttons = []
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Select Titleblock"
        self.Width = 425
        self.Height = 500
        self.MinWidth = 425
        self.MinHeight = 500
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(8)

        for i in range(4):
            row = RowDefinition()
            if i == 2:
                row.Height = GridLength(1, GridUnitType.Star)
            else:
                row.Height = GridLength.Auto
            grid.RowDefinitions.Add(row)

        grid.ColumnDefinitions.Add(ColumnDefinition())

        self.label = WpfLabel()
        self.label.Content = "Search and select titleblock:"
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        self.search_box = WpfTextBox()
        self.search_box.Height = 24
        self.search_box.FontFamily = FontFamily("Arial")
        self.search_box.FontSize = 12
        self.search_box.Margin = Thickness(0, 0, 0, 5)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        self.item_panel = StackPanel()
        scroll_viewer = ScrollViewer()
        scroll_viewer.Content = self.item_panel
        scroll_viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll_viewer.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        button_container = Grid()
        button_panel = UniformGrid()
        button_panel.Columns = 2
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.VerticalAlignment = VerticalAlignment.Bottom
        button_panel.Margin = Thickness(0, 5, 0, 0)

        self.select_button = WpfButton()
        self.select_button.Content = "Select"
        self.select_button.FontFamily = FontFamily("Arial")
        self.select_button.FontSize = 12
        self.select_button.Height = 28
        self.select_button.Margin = Thickness(5, 0, 5, 0)
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.cancel_button = WpfButton()
        self.cancel_button.Content = "Cancel"
        self.cancel_button.FontFamily = FontFamily("Arial")
        self.cancel_button.FontSize = 12
        self.cancel_button.Height = 28
        self.cancel_button.Margin = Thickness(5, 0, 5, 0)
        self.cancel_button.Background = Brushes.LightGray
        self.cancel_button.Click += self.cancel_clicked
        button_panel.Children.Add(self.cancel_button)

        button_container.Children.Add(button_panel)
        Grid.SetRow(button_container, 3)
        grid.Children.Add(button_container)

        self.Content = grid
        self.update_items(self.titleblock_list)

    def update_items(self, titleblocks):
        self.item_panel.Children.Clear()
        self.radio_buttons = []

        for tb in titleblocks:
            family_name = tb.FamilyName
            type_name = tb.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            display_name = "{} - {}".format(family_name, type_name)

            radio = RadioButton()
            radio.Content = display_name.replace("_", "__")
            radio.Tag = tb
            radio.GroupName = "TitleblockGroup"
            radio.FontFamily = FontFamily("Arial")
            radio.FontSize = 12
            radio.Margin = Thickness(2)
            radio.Checked += self.radio_checked

            if self.selected_titleblock and self.selected_titleblock.Id == tb.Id:
                radio.IsChecked = True

            self.item_panel.Children.Add(radio)
            self.radio_buttons.append(radio)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower().strip()
        if search_text:
            filtered = [
                tb for tb in self.titleblock_list
                if search_text in tb.FamilyName.lower()
                or search_text in tb.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString().lower()
            ]
        else:
            filtered = self.titleblock_list
        self.update_items(filtered)

    def radio_checked(self, sender, args):
        self.selected_titleblock = sender.Tag

    def select_clicked(self, sender, args):
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


def get_unique_view_name(doc, base_name):
    existing_names = set()
    for view in fec(doc).OfClass(Autodesk.Revit.DB.View).ToElements():
        try:
            existing_names.add(view.Name)
        except:
            pass

    if base_name not in existing_names:
        return base_name

    i = 1
    while True:
        candidate = "{} ({})".format(base_name, i)
        if candidate not in existing_names:
            return candidate
        i += 1


form = SketchDataForm(lines[0], lines[1])
result = form.ShowDialog()

if result and form.snumber is not None:
    snumber = form.snumber
    sname = form.sname
    unique_view_name = get_unique_view_name(doc, sname)

    sheet_exists = any(sheet.SheetNumber == snumber for sheet in fec(doc).OfClass(Autodesk.Revit.DB.ViewSheet).ToElements())
    if sheet_exists:
        TaskDialog.Show("Error", "Sheet Number {} already exists.".format(snumber))
        sys.exit()

    titleblocks = fec(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
    if not titleblocks:
        TaskDialog.Show("Error", "No title blocks found in the document.")
        sys.exit()

    tb_form = TitleblockSelectionFilter(titleblocks)
    if not tb_form.ShowDialog() or not tb_form.selected_titleblock:
        TaskDialog.Show("Error", "No title block selected.")
        sys.exit()

    selected_titleblock = tb_form.selected_titleblock

    if str(curview.ViewType) == 'FloorPlan':
        pickedBox = uidoc.Selection.PickBox(PickBoxStyle.Directional, "Select area for sketch")
        Maxx = pickedBox.Max.X
        Maxy = pickedBox.Max.Y
        Minx = pickedBox.Min.X
        Miny = pickedBox.Min.Y

        newmaxx = max(Maxx, Minx)
        newmaxy = max(Maxy, Miny)
        newminx = min(Maxx, Minx)
        newminy = min(Maxy, Miny)

        bbox = BoundingBoxXYZ()
        bbox.Max = XYZ(newmaxx, newmaxy, 0)
        bbox.Min = XYZ(newminx, newminy, 0)

        t = Transaction(doc, 'Create Sketch')
        t.Start()

        SHEET = ViewSheet.Create(doc, selected_titleblock.Id)
        SHEET.Name = sname
        SHEET.SheetNumber = snumber

        newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
        getnewview = doc.GetElement(newView)
        getnewview.Name = unique_view_name
        getnewview.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(ElementId.InvalidElementId)
        getnewview.CropBoxActive = True
        getnewview.CropBoxVisible = True
        getnewview.CropBox = bbox

        x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
        y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
        ViewLocation = XYZ(x, y, 0.0)

        NEWSHEET = Viewport.Create(doc, SHEET.Id, newView, ViewLocation)
        t.Commit()
        uidoc.RequestViewChange(SHEET)

    else:
        t = Transaction(doc, 'Create Sketch')
        t.Start()

        SHEET = ViewSheet.Create(doc, selected_titleblock.Id)
        SHEET.Name = sname
        SHEET.SheetNumber = snumber

        newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
        getnewview = doc.GetElement(newView)
        getnewview.Name = unique_view_name
        getnewview.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(ElementId.InvalidElementId)
        getnewview.CropBoxActive = True
        getnewview.CropBoxVisible = True

        x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
        y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
        ViewLocation = XYZ(x, y, 0.0)

        NEWSHEET = Viewport.Create(doc, SHEET.Id, newView, ViewLocation)
        t.Commit()
        uidoc.RequestViewChange(SHEET)

    with open(filepath, 'w') as the_file:
        line1 = snumber + '\n'
        line2 = sname + '\n'
        the_file.writelines([line1, line2])