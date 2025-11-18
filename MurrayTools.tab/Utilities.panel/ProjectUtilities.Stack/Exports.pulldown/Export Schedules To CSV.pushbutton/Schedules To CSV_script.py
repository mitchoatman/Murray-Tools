from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, ViewType
from Autodesk.Revit.UI import TaskDialog

# Add CLR references for WPF + WinForms
import os, re, sys, clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')

import System
from System import IO
from System.Windows import Window, Thickness
from System.Windows.Controls import Label, TextBox, Button, ScrollViewer, StackPanel, Grid, Orientation, CheckBox
from System.Windows.Media import FontFamily
from System.Windows import HorizontalAlignment, GridLength, GridUnitType, ResizeMode
from System.Windows.Controls import ScrollBarVisibility
from System.Windows.Forms import SaveFileDialog, DialogResult

# ------------------------------
# Helpers
# ------------------------------
def cleanup_filename(name):
    """Remove invalid filename characters."""
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def get_schedules(doc):
    """Return all schedules in the doc."""
    collector = FilteredElementCollector(doc).OfClass(DB.ViewSchedule).WhereElementIsNotElementType()
    return sorted(
        [v for v in collector if v.ViewType == ViewType.Schedule and not v.IsTemplate],
        key=lambda x: x.Name
    )


# ------------------------------
# WPF Dialog for Opening Excel
# ------------------------------
class OpenExcelDialog(Window):
    def __init__(self, filepath):
        self.filepath = filepath
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Export Successful"
        self.Width = 320
        self.Height = 150
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(10)
        grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=GridLength.Auto))

        # Success message
        label = Label(Content="Schedule exported successfully.\nWould you like to open the Conversion-Template.xltx?")
        label.FontFamily = FontFamily("Arial")
        label.FontSize = 12
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        # Buttons
        button_panel = StackPanel(Orientation=Orientation.Horizontal,
                                  HorizontalAlignment=HorizontalAlignment.Center,
                                  Margin=Thickness(0, 10, 0, 0))

        open_btn = Button(Content="Open", FontFamily=FontFamily("Arial"), FontSize=12,
                          Margin=Thickness(0, 0, 20, 0), Height=25)
        open_btn.Click += self.open_clicked
        button_panel.Children.Add(open_btn)

        continue_btn = Button(Content="Continue", FontFamily=FontFamily("Arial"), FontSize=12, Height=25)
        continue_btn.Click += self.continue_clicked
        button_panel.Children.Add(continue_btn)

        Grid.SetRow(button_panel, 1)
        grid.Children.Add(button_panel)

        self.Content = grid

    def open_clicked(self, sender, args):
        try:
            System.Diagnostics.Process.Start(self.filepath)
        except Exception as e:
            TaskDialog.Show("Error", "Failed to open file:\n{}".format(str(e)))
        self.Close()

    def continue_clicked(self, sender, args):
        self.Close()


# ------------------------------
# WPF Dialog for Schedule Selection
# ------------------------------
class SelectSchedulesWindow(Window):
    def __init__(self, schedule_list):
        self.selected_schedules = []
        self.schedule_list = sorted(schedule_list, key=lambda x: x.Name)
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Select Schedules to Export"
        self.Width = 400
        self.Height = 400
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(5)
        for i in range(4):  # rows
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Label
        label = Label(Content="Search and select schedules:")
        label.FontFamily = FontFamily("Arial")
        label.FontSize = 16
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        # Search Box
        self.search_box = TextBox(Height=22, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        # Scrollable Checkbox Panel
        self.checkbox_panel = StackPanel(Orientation=Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel)
        scroll_viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.schedule_list)

        # Buttons
        button_panel = StackPanel(Orientation=Orientation.Horizontal,
                                  HorizontalAlignment=HorizontalAlignment.Center,
                                  Margin=Thickness(0, 10, 0, 10))

        select_btn = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25,
                            Margin=Thickness(0, 0, 20, 0))
        select_btn.Click += self.select_clicked
        button_panel.Children.Add(select_btn)

        checkall_btn = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25)
        checkall_btn.Click += self.check_all_clicked
        button_panel.Children.Add(checkall_btn)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        self.Content = grid

    def update_checkboxes(self, schedules):
        prev_selected = set(self.selected_schedules)
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []

        for sched in schedules:
            cb = CheckBox(Content=sched.Name)
            cb.Tag = sched
            cb.Click += self.checkbox_clicked
            if sched in prev_selected:
                cb.IsChecked = True
            self.checkbox_panel.Children.Add(cb)
            self.checkboxes.append(cb)

    def search_changed(self, sender, args):
        text = self.search_box.Text.lower()
        filtered = [s for s in self.schedule_list if text in s.Name.lower()]
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_schedules = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        self.selected_schedules = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def select_clicked(self, sender, args):
        self.selected_schedules = [cb.Tag for cb in self.checkboxes if cb.IsChecked]
        self.DialogResult = True
        self.Close()


# ------------------------------
# Main
# ------------------------------
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
schedules = get_schedules(doc)

if schedules:
    form = SelectSchedulesWindow(schedules)
    if form.ShowDialog() and form.selected_schedules:
        vseop = DB.ViewScheduleExportOptions()
        vseop.ColumnHeaders = DB.ExportColumnHeaders.OneRow
        vseop.FieldDelimiter = ','
        vseop.Title = False
        vseop.HeadersFootersBlanks = True

        script_dir = os.path.dirname(__file__)
        excel_template_path = os.path.join(script_dir, "Conversion-Template.xltx")

        for sched in form.selected_schedules:
            schedule_name = cleanup_filename(sched.Name)

            save_dialog = SaveFileDialog()
            save_dialog.Title = "Save CSV File"
            save_dialog.Filter = "CSV Files (*.csv)|*.csv"
            save_dialog.DefaultExt = "csv"
            save_dialog.InitialDirectory = os.path.expandvars("%USERPROFILE%\\Desktop")
            save_dialog.FileName = schedule_name

            if save_dialog.ShowDialog() == DialogResult.OK:
                folder = os.path.dirname(save_dialog.FileName)
                fname = os.path.basename(save_dialog.FileName)

                sched.Export(folder, fname, vseop)
                
                open_dialog = OpenExcelDialog(excel_template_path)
                open_dialog.ShowDialog()
else:
    TaskDialog.Show("Error", "No schedules found in the document.")
    sys.exit()