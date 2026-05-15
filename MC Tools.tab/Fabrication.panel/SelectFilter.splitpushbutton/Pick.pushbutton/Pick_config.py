import Autodesk
from collections import namedtuple
from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredWorksetCollector, WorksetKind, ElementId
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.UI import TaskDialog
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

import System
from System.Windows import Window, Thickness
from System.Windows.Controls import Label, TextBox, Button, ScrollViewer, StackPanel
from System.Windows.Media import FontFamily
from System.Windows import ResizeMode, HorizontalAlignment
from Autodesk.Revit.Exceptions import OperationCanceledException

WorksetOption = namedtuple('WorksetOption', ['name', 'workset'])


class PickByWorksetSelectionFilter(ISelectionFilter):
    def __init__(self, workset_opts):
        self.workset_ids = set([ws_opt.workset.Id.IntegerValue for ws_opt in workset_opts])

    def AllowElement(self, element):
        try:
            return element.WorksetId.IntegerValue in self.workset_ids
        except:
            return False

    def AllowReference(self, reference, point):
        return False


def get_project_worksets(doc):
    collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    worksets = [ws for ws in collector]
    return sorted(worksets, key=lambda x: x.Name)


def pick_by_worksets(workset_opts, uidoc):
    if not workset_opts:
        return

    ws_filter = PickByWorksetSelectionFilter(workset_opts)

    try:
        selection_list = uidoc.Selection.PickElementsByRectangle(ws_filter)
    except OperationCanceledException:
        return
    except Exception:
        return

    if selection_list:
        filtered_ids = List[ElementId]([e.Id for e in selection_list])
        uidoc.Selection.SetElementIds(filtered_ids)


class FilterSelectionByWorkset(Window):
    def __init__(self, workset_list):
        self.selected_worksets = []
        self.workset_list = sorted(workset_list, key=lambda x: x.Name)
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Filter Selection by Workset"
        self.Width = 400
        self.Height = 435
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        main_panel = StackPanel(Margin=Thickness(8))

        self.label = Label(Content="Search and select worksets:")
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        main_panel.Children.Add(self.label)

        self.search_box = TextBox(Height=22, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        main_panel.Children.Add(self.search_box)

        self.checkbox_panel = StackPanel()
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=System.Windows.Controls.ScrollBarVisibility.Auto)
        scroll_viewer.Height = 280
        scroll_viewer.Margin = Thickness(0, 5, 0, 5)
        main_panel.Children.Add(scroll_viewer)

        self.update_checkboxes(self.workset_list)

        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal,
                                  HorizontalAlignment=HorizontalAlignment.Center,
                                  Margin=Thickness(0, 10, 0, 0))

        self.select_button = Button(Content="Select", Width=60, Height=25, Margin=Thickness(0, 0, 20, 0))
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", Width=75, Height=25)
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        main_panel.Children.Add(button_panel)

        self.Content = main_panel

    def update_checkboxes(self, worksets):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []

        for workset in worksets:
            checkbox = System.Windows.Controls.CheckBox(Content=workset.Name)
            checkbox.Tag = workset
            checkbox.Click += self.checkbox_clicked

            if workset.Name in self.selected_worksets:
                checkbox.IsChecked = True

            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_worksets = [cb.Tag.Name for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        self.selected_worksets = [cb.Tag.Name for cb in self.checkboxes if cb.IsChecked]

    def select_clicked(self, sender, args):
        self.selected_worksets = [cb.Tag.Name for cb in self.checkboxes if cb.IsChecked]
        self.DialogResult = True
        self.Close()

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [w for w in self.workset_list if search_text in w.Name.lower()]
        self.update_checkboxes(filtered)


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

source_worksets = get_project_worksets(doc)

if source_worksets:
    form = FilterSelectionByWorkset(source_worksets)
    if form.ShowDialog():
        workset_opts = [WorksetOption(name=x.Name, workset=x)
                        for x in source_worksets if x.Name in form.selected_worksets]
        pick_by_worksets(workset_opts, uidoc)
else:
    TaskDialog.Show("Error", "No user worksets found in this project.")