# -*- coding: UTF-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, \
    ParameterValueProvider, ElementId, Transaction, FilterStringEquals, \
    FilterStringRule, ElementParameterFilter, LogicalOrFilter, TemporaryViewMode
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString
import clr, sys
from Autodesk.Revit.UI import TaskDialog
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
from System.Windows import Application, Window, Thickness, HorizontalAlignment, \
    VerticalAlignment, ResizeMode, WindowStartupLocation, GridLength, GridUnitType
from System.Windows.Controls import Button, TextBox, CheckBox, Grid, RowDefinition, \
    ColumnDefinition, Label, StackPanel, ScrollViewer, Orientation, ScrollBarVisibility
from System.Windows.Media import Brushes, FontFamily
from System.Windows.Controls.Primitives import UniformGrid

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers).WhereElementIsNotElementType().ToElements()
pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework).WhereElementIsNotElementType().ToElements()
duct_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationDuctwork).WhereElementIsNotElementType().ToElements()

def create_filter_2023_newer(key_parameter, element_value):
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), element_value)
    return ElementParameterFilter(f_rule)

def create_filter_2022_older(key_parameter, element_value):
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    caseSensitive = False
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), element_value, caseSensitive)
    return ElementParameterFilter(f_rule)

class ServiceSelectionForm(object):
    def __init__(self, service_list):
        self.selected_services = []
        self.service_list = sorted(service_list)
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()

    def InitializeComponents(self):
        self._window = Window()
        self._window.Title = "Service Visibility"
        self._window.Width = 400
        self._window.Height = 400
        self._window.MinWidth = self._window.Width
        self._window.MinHeight = self._window.Height
        self._window.ResizeMode = ResizeMode.NoResize
        self._window.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(5)
        grid.VerticalAlignment = VerticalAlignment.Stretch
        grid.HorizontalAlignment = HorizontalAlignment.Stretch

        for i in range(4):
            row = RowDefinition()
            if i == 2:
                row.Height = GridLength(1, GridUnitType.Star)
            else:
                row.Height = GridLength.Auto
            grid.RowDefinitions.Add(row)
        grid.ColumnDefinitions.Add(ColumnDefinition())

        self.label = Label()
        self.label.Content = "Search and select services:"
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        self.search_box = TextBox()
        self.search_box.Height = 20
        self.search_box.Margin = Thickness(0)
        self.search_box.FontFamily = FontFamily("Arial")
        self.search_box.FontSize = 12
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        self.checkbox_panel = StackPanel()
        self.checkbox_panel.Orientation = Orientation.Vertical
        scroll_viewer = ScrollViewer()
        scroll_viewer.Content = self.checkbox_panel
        scroll_viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll_viewer.Margin = Thickness(0, 1, 0, 1)
        scroll_viewer.VerticalAlignment = VerticalAlignment.Stretch
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.service_list)

        # ✅ Anchored button panel at absolute bottom
        button_container = Grid()
        button_container.HorizontalAlignment = HorizontalAlignment.Stretch
        button_container.VerticalAlignment = VerticalAlignment.Bottom

        button_panel = UniformGrid()
        button_panel.Columns = 4
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.VerticalAlignment = VerticalAlignment.Bottom
        button_panel.Margin = Thickness(0, 10, 0, 10)

        # Buttons
        self.reset_button = Button()
        self.reset_button.Content = "Reset View"
        self.reset_button.Background = Brushes.Red
        self.reset_button.FontFamily = FontFamily("Arial")
        self.reset_button.FontSize = 12
        self.reset_button.Height = 25
        self.reset_button.Margin = Thickness(5, 0, 5, 0)  # Gap on left and right
        self.reset_button.Click += self.reset_clicked
        button_panel.Children.Add(self.reset_button)

        self.hide_button = Button()
        self.hide_button.Content = "Hide"
        self.hide_button.FontFamily = FontFamily("Arial")
        self.hide_button.FontSize = 12
        self.hide_button.Height = 25
        self.hide_button.Margin = Thickness(5, 0, 5, 0)  # Gap on left and right
        self.hide_button.Click += self.hide_clicked
        button_panel.Children.Add(self.hide_button)

        self.isolate_button = Button()
        self.isolate_button.Content = "Isolate"
        self.isolate_button.FontFamily = FontFamily("Arial")
        self.isolate_button.FontSize = 12
        self.isolate_button.Height = 25
        self.isolate_button.Margin = Thickness(5, 0, 5, 0)  # Gap on left and right
        self.isolate_button.Click += self.isolate_clicked
        button_panel.Children.Add(self.isolate_button)

        self.check_all_button = Button()
        self.check_all_button.Content = "All / None"
        self.check_all_button.FontFamily = FontFamily("Arial")
        self.check_all_button.FontSize = 12
        self.check_all_button.Height = 25
        self.check_all_button.Margin = Thickness(5, 0, 5, 0)  # Gap on left and right
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        button_container.Children.Add(button_panel)
        Grid.SetRow(button_container, 3)
        grid.Children.Add(button_container)

        self._window.Content = grid
        self._window.SizeChanged += self.on_resize


    def update_checkboxes(self, services):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for service in services:
            checkbox = CheckBox()
            checkbox.Content = service
            checkbox.FontFamily = FontFamily("Arial")
            checkbox.FontSize = 12
            checkbox.Margin = Thickness(2)
            checkbox.IsChecked = service in self.selected_services
            checkbox.Checked += self.checkbox_changed
            checkbox.Unchecked += self.checkbox_changed
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered_services = [s for s in self.service_list if search_text in s.lower()]
        self.update_checkboxes(filtered_services)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for checkbox in self.checkboxes:
            checkbox.IsChecked = self.check_all_state
        self.selected_services = [cb.Content for cb in self.checkboxes if cb.IsChecked]

    def on_resize(self, sender, args):
        pass

    def checkbox_changed(self, sender, args):
        self.selected_services = [cb.Content for cb in self.checkboxes if cb.IsChecked]

    def reset_clicked(self, sender, args):
        try:
            t = Transaction(doc, "Reset Temporary Hide/Isolate")
            t.Start()
            curview.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
            t.Commit()
            self._window.Close()
        except Exception as e:
            print("Reset Error: {}".format(str(e)))

    def hide_clicked(self, sender, args):
        self.selected_services = [cb.Content for cb in self.checkboxes if cb.IsChecked]
        if self.selected_services:
            try:
                filters = []
                for name in self.selected_services:
                    f = create_filter_2023_newer(BuiltInParameter.FABRICATION_SERVICE_NAME, name) if RevitINT > 2022 else create_filter_2022_older(BuiltInParameter.FABRICATION_SERVICE_NAME, name)
                    filters.append(f)
                collector = FilteredElementCollector(doc).WherePasses(LogicalOrFilter(filters)).ToElementIds()
                t = Transaction(doc, "Hide Services")
                t.Start()
                curview.HideElementsTemporary(collector)
                t.Commit()
                self._window.Close()
            except Exception as e:
                print("Hide Error: {}".format(str(e)))

    def isolate_clicked(self, sender, args):
        self.selected_services = [cb.Content for cb in self.checkboxes if cb.IsChecked]
        if self.selected_services:
            try:
                filters = []
                for name in self.selected_services:
                    f = create_filter_2023_newer(BuiltInParameter.FABRICATION_SERVICE_NAME, name) if RevitINT > 2022 else create_filter_2022_older(BuiltInParameter.FABRICATION_SERVICE_NAME, name)
                    filters.append(f)
                collector = FilteredElementCollector(doc).WherePasses(LogicalOrFilter(filters)).ToElementIds()
                t = Transaction(doc, "Isolate Services")
                t.Start()
                curview.IsolateElementsTemporary(collector)
                t.Commit()
                self._window.Close()
            except Exception as e:
                print("Isolate Error: {}".format(str(e)))

    def ShowDialog(self):
        self._window.ShowDialog()

# Collect unique services
SrvcList = []
for item in list(hanger_collector) + list(pipe_collector) + list(duct_collector):
    name = get_parameter_value_by_name_AsString(item, 'Fabrication Service Name')
    if name:
        SrvcList.append(name)

unique_services = set(SrvcList)
if not unique_services:
    TaskDialog.Show("Error", "No fabrication services found in the current view.")
    sys.exit()

form = ServiceSelectionForm(unique_services)
form.ShowDialog()
