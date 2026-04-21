# -*- coding: UTF-8 -*-
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, BuiltInCategory, Transaction, ViewSheet, ViewDuplicateOption, Viewport, ElementId
from Autodesk.Revit.UI.Selection import PickBoxStyle
from Autodesk.Revit.UI import TaskDialog
import sys, os, clr

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
import System
from System.Windows.Forms import Application, Form, FormStartPosition, FormBorderStyle
from System.Drawing import Point, Size, Font, FontStyle
from System.Windows.Controls import Label as WpfLabel, TextBox as WpfTextBox, Button as WpfButton, ScrollViewer, StackPanel, Grid, Orientation, CheckBox
from System.Windows import Window, Thickness, ResizeMode, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType
from System.Windows.Media import FontFamily

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
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
class TXT_Form(Form):
    def __init__(self):
        self.Text = 'Sketch Data'
        self.Size = Size(200, 260)
        self.StartPosition = FormStartPosition.CenterScreen
        self.TopMost = True
        self.ShowIcon = False
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.snumber = None
        self.sname = None
        # Label for Sheet Number
        self.label_textbox = System.Windows.Forms.Label()
        self.label_textbox.Text = 'Sheet Number:'
        self.label_textbox.ForeColor = System.Drawing.Color.Black
        self.label_textbox.Font = Font("Arial", 12)
        self.label_textbox.Location = Point(15, 20)
        self.label_textbox.AutoSize = True
        self.Controls.Add(self.label_textbox)
        # TextBox for Sheet Number
        self.textBox1 = System.Windows.Forms.TextBox()
        self.textBox1.Text = lines[0]
        self.textBox1.Location = Point(15, 50)
        self.textBox1.Size = Size(150, 30)
        self.textBox1.Font = Font("Arial", 10.25, FontStyle.Regular) # Increased from 8.25 to 10.25
        self.Controls.Add(self.textBox1)
        # Label for Sheet Name
        self.label_textbox2 = System.Windows.Forms.Label()
        self.label_textbox2.Text = 'Sheet Name:'
        self.label_textbox2.ForeColor = System.Drawing.Color.Black
        self.label_textbox2.Font = Font("Arial", 12)
        self.label_textbox2.Location = Point(15, 85)
        self.label_textbox2.AutoSize = True
        self.Controls.Add(self.label_textbox2)
        # TextBox for Sheet Name
        self.textBox2 = System.Windows.Forms.TextBox()
        self.textBox2.Text = lines[1]
        self.textBox2.Location = Point(15, 115)
        self.textBox2.Size = Size(150, 30)
        self.textBox2.Font = Font("Arial", 10.25, FontStyle.Regular) # Increased from 8.25 to 10.25
        self.Controls.Add(self.textBox2)
        # Button
        self.button = System.Windows.Forms.Button()
        self.button.Text = 'Set Sketch Data'
        self.button.AutoSize = True
        self.Controls.Add(self.button)
        self.button.Font = Font("Arial", 12)
        self.button.Location = Point((self.ClientSize.Width - self.button.Width) // 2, 165)
        self.button.Height = 25
        self.button.Click += self.on_click
        self.textBox1.Focus()
        self.textBox1.SelectAll()
    def on_click(self, sender, event):
        self.snumber = self.textBox1.Text
        self.sname = self.textBox2.Text
        self.Close()
class TitleblockSelectionFilter(Window):
    def __init__(self, titleblocks):
        self.selected_titleblock = None
        self.titleblock_list = sorted(titleblocks, key=lambda x: x.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()
    def InitializeComponents(self):
        self.Title = "Select Titleblock"
        self.Width = 400
        self.Height = 400
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        grid = Grid()
        grid.Margin = Thickness(5)
        for i in range(4): # rows for: label, search box, scroll, buttons
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())
        # Row 0 - Label
        self.label = WpfLabel(Content="Select a titleblock:")
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)
        # Row 1 - Search Box
        self.search_box = WpfTextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)
        # Row 2 - Scrollable Checkbox Panel
        self.checkbox_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=System.Windows.Controls.ScrollBarVisibility.Auto)
        scroll_viewer.Margin = Thickness(0, 1, 0, 1)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)
        self.update_checkboxes(self.titleblock_list)
        # Row 3 - Button Panel
        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 10, 0, 10))
        self.select_button = WpfButton(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=50, HorizontalAlignment=HorizontalAlignment.Center)
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)
        self.check_all_button = WpfButton(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=70, HorizontalAlignment=HorizontalAlignment.Center)
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)
        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)
        # Set window content
        self.Content = grid
        self.SizeChanged += self.on_resize
    def update_checkboxes(self, titleblocks):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for tb in titleblocks:
            family_name = tb.FamilyName
            type_name = tb.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            display_name = "{} - {}".format(family_name, type_name)
            checkbox = CheckBox(Content=display_name)
            checkbox.Tag = tb
            checkbox.Click += self.checkbox_clicked
            if self.selected_titleblock and type_name == self.selected_titleblock.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString():
                checkbox.IsChecked = True
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)
    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [tb for tb in self.titleblock_list if search_text in tb.FamilyName.lower() or search_text in tb.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString().lower()]
        self.update_checkboxes(filtered)
    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_titleblock = [cb.Tag for cb in self.checkboxes if cb.IsChecked][-1] if self.check_all_state else None
    def checkbox_clicked(self, sender, args):
        checked = [cb for cb in self.checkboxes if cb.IsChecked]
        if checked:
            self.selected_titleblock = checked[-1].Tag
            for cb in self.checkboxes:
                if cb != checked[-1]:
                    cb.IsChecked = False
    def select_clicked(self, sender, args):
        checked = [cb for cb in self.checkboxes if cb.IsChecked]
        if checked:
            self.selected_titleblock = checked[-1].Tag
        self.DialogResult = True
        self.Close()
    def on_resize(self, sender, args):
        pass
form = TXT_Form()
Application.Run(form)
if form.snumber is not None:
    snumber = form.snumber
    sname = form.sname
    # Check if sheet already exists
    sheet_exists = any(sheet.SheetNumber == snumber for sheet in fec(doc).OfClass(Autodesk.Revit.DB.ViewSheet).ToElements())
    if sheet_exists:
        TaskDialog.Show("Error", "Sheet Name or Number {} already exists.".format(snumber))
        sys.exit()
    # Check if view already exists
    view_exists = any(view.Name == sname for view in fec(doc).OfClass(Autodesk.Revit.DB.View).ToElements())
    if view_exists:
        TaskDialog.Show("Error", "View with name {} already exists.".format(sname))
        sys.exit()
    # Get all titleblocks
    titleblocks = fec(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
    if not titleblocks:
        TaskDialog.Show("Error", "No title blocks found in the document.")
        sys.exit()
    # Show titleblock selection dialog
    tb_form = TitleblockSelectionFilter(titleblocks)
    if not tb_form.ShowDialog() or not tb_form.selected_titleblock:
        TaskDialog.Show("Error", "No title block selected.")
        sys.exit()
    selected_titleblock = tb_form.selected_titleblock
    if str(curview.ViewType) == 'FloorPlan':
        # Prompt user for box and make sure mins are mins and maxs are maxs
        pickedBox = uidoc.Selection.PickBox(PickBoxStyle.Directional, "Select area for sketch")
        Maxx = pickedBox.Max.X
        Maxy = pickedBox.Max.Y
        Minx = pickedBox.Min.X
        Miny = pickedBox.Min.Y
        newmaxx = max(Maxx, Minx)
        newmaxy = max(Maxy, Miny)
        newminx = min(Maxx, Minx)
        newminy = min(Maxy, Miny)
        # Make bounding box of the points selected
        bbox = BoundingBoxXYZ()
        bbox.Max = XYZ(newmaxx, newmaxy, 0)
        bbox.Min = XYZ(newminx, newminy, 0)
        # Define a transaction variable and describe the transaction
        t = Transaction(doc, 'Create Sketch')
        # Begin new transaction
        t.Start()
        SHEET = ViewSheet.Create(doc, selected_titleblock.Id)
        SHEET.Name = sname
        SHEET.SheetNumber = snumber
        newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
        getnewview = doc.GetElement(newView)
        getnewview.Name = sname
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
        # Define a transaction variable and describe the transaction
        t = Transaction(doc, 'Create Sketch')
        # Begin new transaction
        t.Start()
        SHEET = ViewSheet.Create(doc, selected_titleblock.Id)
        SHEET.Name = sname
        SHEET.SheetNumber = snumber
        newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
        getnewview = doc.GetElement(newView)
        getnewview.Name = sname
        getnewview.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(ElementId.InvalidElementId)
        getnewview.CropBoxActive = True
        getnewview.CropBoxVisible = True
        x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
        y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
        ViewLocation = XYZ(x, y, 0.0)
        NEWSHEET = Viewport.Create(doc, SHEET.Id, newView, ViewLocation)
        t.Commit()
        uidoc.RequestViewChange(SHEET)
    # Update the text file with the new values
    with open(filepath, 'w') as the_file:
        line1 = snumber + '\n'
        line2 = sname + '\n'
        the_file.writelines([line1, line2])