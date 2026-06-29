# -*- coding: utf-8 -*-
import os
import re
import clr
import System

clr.AddReference("System")
clr.AddReference("System.Core")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from pyrevit import forms
import Autodesk.Revit.UI as UI

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    Transaction,
    FilteredElementCollector,
    ElementId,
    FabricationPart,
    ParameterFilterRuleFactory,
    Color,
    BuiltInCategory,
    ParameterFilterElement,
    ElementParameterFilter,
    OverrideGraphicSettings,
    View
)
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from random import randint

try:
    from Parameters.Add_SharedParameters import Shared_Params
except Exception:
    def Shared_Params():
        pass

try:
    from Parameters.Get_Set_Params import set_parameter_by_name
except Exception:
    def set_parameter_by_name(elem, param_name, value):
        param = elem.LookupParameter(param_name)
        if param and not param.IsReadOnly:
            param.Set(value)


PANE_XAML = """
<Page
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Background="#FFF5F5F5">

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,8">
            <TextBlock Text="New Line Number:"
                       Margin="0,0,0,4"/>
            <TextBox x:Name="line_input_tb"
                     Height="24"/>
            <Button x:Name="apply_btn"
                    Content="Apply to Selected"
                    Height="26"
                    Margin="0,6,0,0"/>
        </StackPanel>

        <StackPanel Grid.Row="1" Margin="0,0,0,8">
            <TextBlock Text="Search:"
                       Margin="0,0,0,4"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="32"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="search_tb"
                         Grid.Column="0"
                         Height="24"/>
                <Button x:Name="clear_search_btn"
                        Grid.Column="1"
                        Content="X"
                        Height="24"
                        Margin="6,0,0,0"/>
            </Grid>
        </StackPanel>

        <TextBlock Grid.Row="2"
                   Text="Line Numbers in Project"
                   FontWeight="SemiBold"
                   Margin="0,0,0,6"/>

        <ListBox x:Name="line_numbers_lb"
                 Grid.Row="3"
                 MinHeight="160"
                 BorderBrush="#FFD0D0D0"
                 BorderThickness="1"/>

        <Button x:Name="make_filters_btn"
                Grid.Row="4"
                Content="Make View Filters"
                Height="26"
                Margin="0,8,0,0"/>

        <TextBlock x:Name="status_tb"
                   Grid.Row="5"
                   Margin="0,8,0,0"
                   Foreground="#666666"
                   TextWrapping="Wrap"
                   Text="Ready."/>
    </Grid>
</Page>
"""


PARAM_NAME = "FP_Line Number"
FILTER_PREFIX = "LINE - "
FOLDER_NAME = r"C:\Temp"
FILE_PATH = os.path.join(FOLDER_NAME, "Ribbon_LineNumber.txt")

state = UI.DockablePaneState()
state.DockPosition = UI.DockPosition.Right


def natural_key(value):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r"([0-9]+)", value or "")]


def ensure_input_file():
    if not os.path.exists(FOLDER_NAME):
        os.makedirs(FOLDER_NAME)

    if not os.path.exists(FILE_PATH):
        with open(FILE_PATH, "w") as f:
            f.write("123")


def read_previous_input():
    ensure_input_file()
    try:
        with open(FILE_PATH, "r") as f:
            return f.read().strip()
    except Exception:
        return ""


def write_previous_input(value):
    ensure_input_file()
    with open(FILE_PATH, "w") as f:
        f.write(value or "")


def get_line_numbers_in_project(doc):
    values = set()

    collector = (
        FilteredElementCollector(doc)
        .OfClass(FabricationPart)
        .WhereElementIsNotElementType()
    )

    for elem in collector:
        try:
            param = elem.LookupParameter(PARAM_NAME)
            if param and param.HasValue:
                val = param.AsString()
                if val and val.strip():
                    values.add(val.strip())
        except Exception:
            pass

    return sorted(values, key=natural_key)


def get_elements_by_line_number(doc, line_number):
    matches = []

    collector = (
        FilteredElementCollector(doc)
        .OfClass(FabricationPart)
        .WhereElementIsNotElementType()
    )

    for elem in collector:
        try:
            param = elem.LookupParameter(PARAM_NAME)
            if param and param.HasValue and param.AsString() == line_number:
                matches.append(elem)
        except Exception:
            pass

    return matches


def show_elements(uidoc, elements):
    if not elements:
        return

    ids = List[ElementId]()
    for elem in elements:
        ids.Add(elem.Id)

    uidoc.Selection.SetElementIds(ids)
    uidoc.ShowElements(ids)


def try_set_customdata(elem, value):
    try:
        elem.SetPartCustomDataText(1, value)
    except Exception:
        pass


def random_color():
    return randint(0, 230), randint(0, 230), randint(0, 230)


def get_target_view_or_template(doc):
    curview = doc.ActiveView
    view_template_id = curview.ViewTemplateId
    if not view_template_id.Equals(ElementId.InvalidElementId):
        return doc.GetElement(view_template_id)
    return curview


def get_line_number_param_id(doc):
    sample_element = (
        FilteredElementCollector(doc)
        .OfClass(FabricationPart)
        .WhereElementIsNotElementType()
        .FirstElement()
    )

    if not sample_element:
        return None

    for p in sample_element.Parameters:
        try:
            if p.Definition.Name == PARAM_NAME:
                return p.Id
        except Exception:
            pass

    return None


def get_filter_categories():
    categories = List[ElementId]()
    categories.Add(ElementId(BuiltInCategory.OST_FabricationHangers))
    categories.Add(ElementId(BuiltInCategory.OST_FabricationPipework))
    categories.Add(ElementId(BuiltInCategory.OST_FabricationDuctwork))
    return categories


class _LineNumberRequestHandler(IExternalEventHandler):
    def __init__(self):
        self.request_name = None
        self.line_number = None
        self.line_numbers = None
        self.pane = None

    def Execute(self, uiapp):
        if self.pane is None:
            return

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            self.pane.set_status("No active document.")
            self.pane.set_all_line_numbers([])
            return

        doc = uidoc.Document

        try:
            if self.request_name == "refresh":
                self._do_refresh(doc)

            elif self.request_name == "show":
                self._do_show(uidoc, doc, self.line_number)

            elif self.request_name == "apply":
                self._do_apply(uidoc, doc, self.line_number)

            elif self.request_name == "make_filters":
                self._do_make_filters(doc, self.line_numbers or [])

        except Exception as ex:
            self.pane.set_status("Error: {}".format(str(ex)))

        finally:
            self.request_name = None
            self.line_number = None
            self.line_numbers = None

    def GetName(self):
        return "Line Number Pane External Event"

    def _do_refresh(self, doc):
        values = get_line_numbers_in_project(doc)
        self.pane.set_all_line_numbers(values)
        self.pane.set_status("{} line number(s) found in project.".format(len(values)))

    def _do_show(self, uidoc, doc, line_number):
        if not line_number:
            self.pane.set_status("Select a line number first.")
            return

        matches = get_elements_by_line_number(doc, line_number)
        if not matches:
            self.pane.set_status("No elements found for '{}'.".format(line_number))
            return

        show_elements(uidoc, matches)
        self.pane.set_status("Showing {} element(s) for '{}'.".format(len(matches), line_number))

    def _do_apply(self, uidoc, doc, line_number):
        line_number = (line_number or "").strip()
        if not line_number:
            self.pane.set_status("Enter a line number first.")
            return

        selected_ids = list(uidoc.Selection.GetElementIds())
        if not selected_ids:
            self.pane.set_status("Select one or more elements in Revit, then click Apply.")
            return

        Shared_Params()

        changed = 0
        t = Transaction(doc, "Set Line Number")
        t.Start()
        try:
            for eid in selected_ids:
                elem = doc.GetElement(eid)
                if elem is None:
                    continue

                param = elem.LookupParameter(PARAM_NAME)
                if param and not param.IsReadOnly:
                    try:
                        set_parameter_by_name(elem, PARAM_NAME, line_number)
                    except Exception:
                        param.Set(line_number)

                    try_set_customdata(elem, line_number)
                    changed += 1

            t.Commit()
        except Exception:
            if t.HasStarted():
                t.RollBack()
            raise

        write_previous_input(line_number)
        values = get_line_numbers_in_project(doc)
        self.pane.set_all_line_numbers(values)
        self.pane.set_status("Applied '{}' to {} element(s).".format(line_number, changed))

    def _do_make_filters(self, doc, line_numbers):
        line_numbers = [x.strip() for x in line_numbers if x and x.strip()]
        line_numbers = sorted(set(line_numbers), key=natural_key)

        if not line_numbers:
            self.pane.set_status("No line numbers in the list.")
            return

        Shared_Params()

        app = doc.Application
        revit_int = int(app.VersionNumber)
        view_to_modify = get_target_view_or_template(doc)
        param_id = get_line_number_param_id(doc)

        if not param_id:
            self.pane.set_status("Could not find parameter id for '{}'.".format(PARAM_NAME))
            return

        categories = get_filter_categories()

        existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
        existing_filter_names = {f.Name for f in existing_filters}
        existing_filter_dict = {f.Name: f.Id for f in existing_filters}

        try:
            applied_filters = {doc.GetElement(fid).Name: fid for fid in view_to_modify.GetFilters()}
        except Exception:
            applied_filters = {}

        line_filter_color_dict = {}
        views_collector = FilteredElementCollector(doc).OfClass(View)

        for view_elem in views_collector:
            if view_elem.Id.Equals(view_to_modify.Id):
                continue

            try:
                for f_id in view_elem.GetFilters():
                    f_elem = doc.GetElement(f_id)
                    if f_elem and f_elem.Name.startswith(FILTER_PREFIX):
                        if f_elem.Name not in line_filter_color_dict:
                            try:
                                ovr = view_elem.GetFilterOverrides(f_id)
                                col = ovr.ProjectionLineColor
                                line_filter_color_dict[f_elem.Name] = (col.Red, col.Green, col.Blue)
                            except Exception:
                                pass
            except Exception:
                pass

        created_count = 0
        applied_count = 0
        kept_existing_count = 0

        t = Transaction(doc, "Create and Apply Line Number Filters")
        t.Start()
        try:
            for line_number in line_numbers:
                filter_name = FILTER_PREFIX + line_number
                filter_id = None
                keep_existing_overrides = False
                rgb = None

                if filter_name in existing_filter_names:
                    filter_id = existing_filter_dict[filter_name]

                    if filter_name in applied_filters:
                        keep_existing_overrides = True
                        kept_existing_count += 1
                    else:
                        if filter_name in line_filter_color_dict:
                            rgb = line_filter_color_dict[filter_name]
                        else:
                            rgb = random_color()

                else:
                    try:
                        if revit_int < 2023:
                            rule = ParameterFilterRuleFactory.CreateEqualsRule(param_id, line_number, False)
                        else:
                            rule = ParameterFilterRuleFactory.CreateEqualsRule(param_id, line_number)

                        elem_filter = ElementParameterFilter(rule)
                        param_filter = ParameterFilterElement.Create(doc, filter_name, categories)
                        param_filter.SetElementFilter(elem_filter)
                        filter_id = param_filter.Id
                        created_count += 1

                        if filter_name in line_filter_color_dict:
                            rgb = line_filter_color_dict[filter_name]
                        else:
                            rgb = random_color()

                    except Exception:
                        continue

                if filter_name not in applied_filters:
                    try:
                        view_to_modify.AddFilter(filter_id)
                        view_to_modify.SetFilterVisibility(filter_id, True)
                        applied_count += 1
                    except Exception:
                        continue

                if not keep_existing_overrides and rgb is not None:
                    try:
                        overrides = OverrideGraphicSettings()
                        overrides.SetProjectionLineColor(Color(rgb[0], rgb[1], rgb[2]))
                        view_to_modify.SetFilterOverrides(filter_id, overrides)
                    except Exception:
                        pass

            t.Commit()
        except Exception:
            if t.HasStarted():
                t.RollBack()
            raise

        self.pane.set_status(
            "Filters done. Created: {} | Applied: {} | Kept overrides: {}.".format(
                created_count, applied_count, kept_existing_count
            )
        )


class MyPane(forms.WPFPanel):
    panel_id = "7C3D8C2E-7A0D-4A6D-9B3C-6E4A1E9F51A2"
    panel_title = "Line Number"
    panel_source = "inline"
    initial_state = state

    def __init__(self):
        self.load_xaml(PANE_XAML, literal_string=True)

        self._handler = _LineNumberRequestHandler()
        self._handler.pane = self
        self._ext_event = ExternalEvent.Create(self._handler)

        self._all_line_numbers = []

        self.line_input_tb.Text = read_previous_input()

        self.apply_btn.Click += self.on_apply_clicked
        self.make_filters_btn.Click += self.on_make_filters_clicked
        self.clear_search_btn.Click += self.on_clear_search_clicked
        self.search_tb.TextChanged += self.on_search_changed
        self.line_numbers_lb.SelectionChanged += self.on_line_selected
        self.line_numbers_lb.MouseDoubleClick += self.on_line_double_clicked
        self.Loaded += self.on_loaded

    def on_loaded(self, sender, args):
        self.request_refresh()

    def on_line_selected(self, sender, args):
        selected = self.line_numbers_lb.SelectedItem
        if selected:
            self.line_input_tb.Text = str(selected)

    def on_line_double_clicked(self, sender, args):
        selected = self.line_numbers_lb.SelectedItem
        if not selected:
            self.set_status("Select a line number first.")
            return
        self.request_show(str(selected))

    def on_apply_clicked(self, sender, args):
        self.request_apply(self.line_input_tb.Text)

    def on_make_filters_clicked(self, sender, args):
        values = []
        for item in self.line_numbers_lb.Items:
            values.append(str(item))
        self.request_make_filters(values)

    def on_search_changed(self, sender, args):
        self.apply_search_filter()

    def on_clear_search_clicked(self, sender, args):
        self.search_tb.Text = ""

    def request_refresh(self):
        self._handler.request_name = "refresh"
        self._handler.line_number = None
        self._handler.line_numbers = None
        self._ext_event.Raise()

    def request_show(self, line_number):
        self._handler.request_name = "show"
        self._handler.line_number = line_number
        self._handler.line_numbers = None
        self._ext_event.Raise()

    def request_apply(self, line_number):
        self._handler.request_name = "apply"
        self._handler.line_number = line_number
        self._handler.line_numbers = None
        self._ext_event.Raise()

    def request_make_filters(self, line_numbers):
        self._handler.request_name = "make_filters"
        self._handler.line_number = None
        self._handler.line_numbers = list(line_numbers)
        self._ext_event.Raise()

    def set_all_line_numbers(self, values):
        self._all_line_numbers = list(values)
        self.apply_search_filter()

    def apply_search_filter(self):
        current_text = self.line_input_tb.Text
        search_text = (self.search_tb.Text or "").strip().lower()

        if search_text:
            filtered = [x for x in self._all_line_numbers if search_text in x.lower()]
        else:
            filtered = list(self._all_line_numbers)

        self.line_numbers_lb.Items.Clear()
        for value in filtered:
            self.line_numbers_lb.Items.Add(value)

        self.line_input_tb.Text = current_text

    def set_status(self, message):
        self.status_tb.Text = message or ""