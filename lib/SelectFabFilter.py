# -*- coding: utf-8 -*-
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Collections')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Windows import (
    Window, WindowStartupLocation, WindowStyle, GridLength,
    HorizontalAlignment, VerticalAlignment, GridUnitType, Thickness
)
from System.Windows.Controls import (
    Label, ComboBox, Button, ListBox, CheckBox, TextBox, Grid,
    RowDefinition, ColumnDefinition, SelectionMode, StackPanel,
    Orientation, ListBoxItem
)
from System.Windows.Media import Brushes
from System.Collections.Generic import List
from System.Windows.Threading import DispatcherFrame, Dispatcher

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FabricationConfiguration,
    Transaction,
    TemporaryViewMode
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons

from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import (
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsValueString,
    get_parameter_value_by_name_AsInteger
)


class ValueItem(object):
    def __init__(self, property, value):
        self.Property = property
        self.Value = value

    def __str__(self):
        return u"{}: {}".format(self.Property, self.Value)


class RemoveFilterDialog(Window):
    def __init__(self, filter_options, filter_keys):
        self.filter_options = filter_options
        self.filter_keys = filter_keys
        self.selected_filter = None
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Remove Filter"
        self.Width = 300
        self.Height = 140
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = 0
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        grid = Grid()
        self.Content = grid

        for _ in range(2):
            grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

        self.label = Label(Content="Select filter to remove:", Margin=Thickness(10, 5, 0, 0))
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        self.filter_combo = ComboBox(Margin=Thickness(10, 30, 10, 0), Width=260, Height=20)
        for option in self.filter_options:
            self.filter_combo.Items.Add(option)
        if self.filter_options:
            self.filter_combo.SelectedIndex = 0
        Grid.SetRow(self.filter_combo, 0)
        grid.Children.Add(self.filter_combo)

        panel = StackPanel(
            Orientation=Orientation.Horizontal,
            HorizontalAlignment=HorizontalAlignment.Center,
            Margin=Thickness(0, 10, 0, 10)
        )
        Grid.SetRow(panel, 1)
        grid.Children.Add(panel)

        btn = Button(Content="Remove", Width=80, Height=25, Margin=Thickness(0, 0, 5, 0))
        btn.Click += self.ok_clicked
        panel.Children.Add(btn)

        btn = Button(Content="Remove All", Width=80, Height=25, Margin=Thickness(5, 0, 5, 0))
        btn.Click += self.remove_all_clicked
        panel.Children.Add(btn)

        btn = Button(Content="Cancel", Width=80, Height=25, Margin=Thickness(5, 0, 0, 0))
        btn.Click += self.cancel_clicked
        panel.Children.Add(btn)

    def ok_clicked(self, s, a):
        if self.filter_combo.SelectedIndex >= 0:
            self.selected_filter = self.filter_keys[self.filter_combo.SelectedIndex]
        self.DialogResult = True
        self.Close()

    def remove_all_clicked(self, s, a):
        self.selected_filter = None
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, s, a):
        self.DialogResult = False
        self.Close()


class MultiPropertyFilterForm(Window):
    def __init__(self, doc, uidoc, curview, config, property_options, fab_elements, all_elements):
        self.doc = doc
        self.uidoc = uidoc
        self.curview = curview
        self.config = config
        self.property_options = property_options
        self.fab_elements = fab_elements
        self.all_elements = all_elements
        self.selected_filters = {}

        self.InitializeComponents()

        if self.property_options:
            self.property_combo.SelectedItem = list(self.property_options.keys())[0]
            self.update_values_list(None, None)

        self.update_filter_display()

    def InitializeComponents(self):
        self.Title = "Multi-Property Filter"
        self.Width = 550
        self.Height = 620
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = 0
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        grid = Grid()
        self.Content = grid

        heights = [
            GridLength(40),
            GridLength.Auto,
            GridLength.Auto,
            GridLength(260),
            GridLength.Auto,
            GridLength(140),
            GridLength.Auto
        ]
        for h in heights:
            grid.RowDefinitions.Add(RowDefinition(Height=h))

        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength.Auto))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))

        self.property_label = Label(
            Content="Select Property:",
            Margin=Thickness(10, 0, 10, 0),
            HorizontalAlignment=HorizontalAlignment.Left
        )
        Grid.SetRow(self.property_label, 0)
        Grid.SetColumn(self.property_label, 1)
        grid.Children.Add(self.property_label)

        self.property_combo = ComboBox(
            Width=160,
            Height=22,
            Margin=Thickness(120, 0, 10, 0),
            HorizontalAlignment=HorizontalAlignment.Left
        )
        for prop in sorted(self.property_options.keys()):
            self.property_combo.Items.Add(prop)
        self.property_combo.SelectionChanged += self.update_values_list
        Grid.SetRow(self.property_combo, 0)
        Grid.SetColumn(self.property_combo, 1)
        grid.Children.Add(self.property_combo)

        self.search_label = Label(
            Content="Search:",
            Margin=Thickness(10, 0, 10, 0),
            HorizontalAlignment=HorizontalAlignment.Left
        )
        Grid.SetRow(self.search_label, 1)
        Grid.SetColumn(self.search_label, 1)
        grid.Children.Add(self.search_label)

        self.search_box = TextBox(
            Width=300,
            Height=20,
            Margin=Thickness(75, 0, 10, 0),
            HorizontalAlignment=HorizontalAlignment.Left
        )
        self.search_box.TextChanged += self.update_values_list
        Grid.SetRow(self.search_box, 1)
        Grid.SetColumn(self.search_box, 1)
        grid.Children.Add(self.search_box)
        self.search_box.Focus()

        self.values_label = Label(
            Content="Select Values:",
            Margin=Thickness(10, 25, 10, 0),
            HorizontalAlignment=HorizontalAlignment.Left
        )
        Grid.SetRow(self.values_label, 2)
        Grid.SetColumn(self.values_label, 1)
        grid.Children.Add(self.values_label)

        self.values_list = ListBox(
            Width=500,
            Height=250,
            Margin=Thickness(10, 0, 10, 0),
            SelectionMode=SelectionMode.Extended
        )
        self.values_list.MouseDoubleClick += self.add_filter
        Grid.SetRow(self.values_list, 3)
        Grid.SetColumn(self.values_list, 1)
        grid.Children.Add(self.values_list)

        self.add_button = Button(
            Content="Add Filter",
            Width=80,
            Height=25,
            Margin=Thickness(10, 20, 10, 0),
            HorizontalAlignment=HorizontalAlignment.Right
        )
        self.add_button.Click += self.add_filter
        Grid.SetRow(self.add_button, 2)
        Grid.SetColumn(self.add_button, 1)
        grid.Children.Add(self.add_button)

        self.filter_button = Button(
            Width=500,
            Height=130,
            Margin=Thickness(10, 10, 10, 0),
            HorizontalContentAlignment=HorizontalAlignment.Left,
            VerticalContentAlignment=VerticalAlignment.Top,
            HorizontalAlignment=HorizontalAlignment.Center
        )
        self.filter_button.ToolTip = "Click here to modify Filters"
        self.filter_button.Click += self.remove_filter
        Grid.SetRow(self.filter_button, 5)
        Grid.SetColumn(self.filter_button, 1)
        grid.Children.Add(self.filter_button)

        self.logic_check = CheckBox(
            Content="AND logic (unchecked = OR)",
            Width=230,
            Margin=Thickness(10, 5, 10, 0),
            HorizontalAlignment=HorizontalAlignment.Center
        )
        Grid.SetRow(self.logic_check, 4)
        Grid.SetColumn(self.logic_check, 1)
        grid.Children.Add(self.logic_check)

        panel = StackPanel(
            Orientation=Orientation.Horizontal,
            HorizontalAlignment=HorizontalAlignment.Center,
            Margin=Thickness(0, 15, 0, 0)
        )
        Grid.SetRow(panel, 6)
        Grid.SetColumn(panel, 1)
        grid.Children.Add(panel)

        for txt, handler in [
            ("Update Data", self.reset_filter_clicked),
            ("Reset View", self.reset_clicked),
            ("Isolate", self.isolate_clicked),
            ("Select", self.select_clicked),
            ("Close", self.cancel_clicked)
        ]:
            btn = Button(
                Content=txt,
                Width=80,
                Height=25,
                Margin=Thickness(0 if txt == "Update Data" else 5, 0, 0 if txt == "Close" else 5, 0)
            )
            if txt == "Reset View":
                btn.Background = Brushes.Red
                btn.Foreground = Brushes.Black
            btn.Click += handler
            panel.Children.Add(btn)

    def exit_frame(self, s, e):
        self.frame.Continue = False

    def reset_filter_clicked(self, sender, args):
        try:
            pre = [self.doc.GetElement(i) for i in self.uidoc.Selection.GetElementIds()]
            self.fab_elements = pre or FilteredElementCollector(
                self.doc, self.curview.Id
            ).OfClass(DB.FabricationPart).WhereElementIsNotElementType().ToElements()

            self.all_elements = pre or FilteredElementCollector(
                self.doc, self.curview.Id
            ).WhereElementIsNotElementType().ToElements()

            self.property_options = {}

            fab_props = [
                'CID', 'ServiceType', 'Service Name', 'Service Abbreviation', 'Size',
                'STRATUS Assembly', 'Line Number', 'STRATUS Status', 'Reference Level',
                'Item Number', 'Bundle Number', 'REF BS Designation', 'REF Line Number',
                'Specification', 'Hanger Rod Size', 'Valve Number', 'Beam Hanger',
                'Product Entry', 'TS_Point_Number', 'TS_Point_Description', 'Alias'
            ]

            for prop in fab_props:
                vals = set(filter(None, [get_property_value(e, prop, self.config, False) for e in self.fab_elements]))
                if vals:
                    self.property_options[prop] = sorted(vals)

            all_props = ['Name', 'Comments', 'Category', 'TS_Point_Number', 'TS_Point_Description']
            for prop in all_props:
                vals = set(filter(None, [get_property_value(e, prop, self.config, False) for e in self.all_elements]))
                if vals:
                    self.property_options[prop] = sorted(vals)

            self.property_combo.Items.Clear()
            for p in sorted(self.property_options.keys()):
                self.property_combo.Items.Add(p)

            if self.property_options:
                first = sorted(self.property_options.keys())[0]
                self.property_combo.SelectedItem = first
                self.search_box.Text = ""
                self.values_list.Items.Clear()
                for v in self.property_options[first]:
                    item = ListBoxItem(Content=str(ValueItem(first, v)))
                    item.Tag = ValueItem(first, v)
                    self.values_list.Items.Add(item)

        except Exception as e:
            dlg = TaskDialog("Error")
            dlg.MainInstruction = "Update Data Error: {0}".format(e)
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()

    def update_values_list(self, sender, args):
        term = self.search_box.Text.lower()
        sel = self.property_combo.SelectedItem
        self.values_list.Items.Clear()

        if term:
            results = []
            for prop, vals in self.property_options.items():
                for v in vals:
                    if term in str(v).lower():
                        results.append(ValueItem(prop, v))
            results.sort(key=lambda x: (x.Property, x.Value))
            for vi in results:
                item = ListBoxItem(Content=str(vi))
                item.Tag = vi
                self.values_list.Items.Add(item)
        elif sel:
            for v in self.property_options.get(sel, []):
                vi = ValueItem(sel, v)
                item = ListBoxItem(Content=str(vi))
                item.Tag = vi
                self.values_list.Items.Add(item)

    def add_filter(self, sender, args):
        sels = self.values_list.SelectedItems
        if not sels:
            dlg = TaskDialog("Warning")
            dlg.MainInstruction = "Please select at least one value."
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()
            return

        from collections import defaultdict
        m = defaultdict(list)
        for it in sels:
            vi = it.Tag
            m[vi.Property].append(vi.Value)

        for prop, vals in m.items():
            self.selected_filters.setdefault(prop, []).append((vals, self.logic_check.IsChecked))

        self.update_filter_display()

    def remove_filter(self, sender, args):
        if not self.selected_filters:
            return

        opts, keys = [], []
        for prop, flist in self.selected_filters.items():
            for vals, andf in flist:
                mode = "AND" if andf else "OR"
                opts.append("{0} ({1}): {2}".format(prop, mode, ", ".join(str(v) for v in vals)))
                keys.append((prop, vals))

        if not opts:
            dlg = TaskDialog("Information")
            dlg.MainInstruction = "No filters to remove."
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()
            return

        dlg = RemoveFilterDialog(opts, keys)
        if dlg.ShowDialog():
            if dlg.selected_filter is None:
                self.selected_filters.clear()
            else:
                p, vals = dlg.selected_filter
                for i, (v, _) in enumerate(self.selected_filters[p]):
                    if v == vals:
                        del self.selected_filters[p][i]
                        break
                if not self.selected_filters[p]:
                    del self.selected_filters[p]

            self.update_filter_display()

    def update_filter_display(self, sender=None, args=None):
        if not self.selected_filters:
            self.filter_button.Content = "No Filters Yet..."
        else:
            txt = "Filters (Click here to modify):\n"
            for prop, fl in self.selected_filters.items():
                txt += "{0}:\n ".format(prop)
                conds = []
                for vals, andf in fl:
                    mode = "AND" if andf else "OR"
                    conds.append("[{0}: {1}]".format(mode, ", ".join(str(v) for v in vals)))
                txt += " ".join(conds) + "\n"
            self.filter_button.Content = txt.strip()

    def reset_clicked(self, sender, args):
        try:
            t = Transaction(self.doc, "Reset Temporary Hide/Isolate")
            t.Start()
            self.curview.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
            t.Commit()
        except Exception as e:
            dlg = TaskDialog("Error")
            dlg.MainInstruction = "Reset Error: {0}".format(e)
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()

    def isolate_clicked(self, sender, args):
        if not self.selected_filters:
            dlg = TaskDialog("Warning")
            dlg.MainInstruction = "No filters selected to isolate."
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()
            return

        try:
            pre = [self.doc.GetElement(i) for i in self.uidoc.Selection.GetElementIds()]
            use_all = any(k in self.selected_filters for k in ("Name", "Comments", "Category", "TS_Point_Number", "TS_Point_Description"))
            elems = pre or (self.all_elements if use_all else self.fab_elements)

            ids = []
            for e in elems:
                if not e or not e.IsValidObject:
                    continue

                ev = {p: get_property_value(e, p, self.config, False) for p in self.selected_filters}
                ok = True

                for p, fl in self.selected_filters.items():
                    and_hits = []
                    or_hits = []
                    for vals, andf in fl:
                        hit = (str(ev[p]) in [str(v) for v in vals])
                        if andf:
                            and_hits.append(hit)
                        else:
                            or_hits.append(hit)

                    if (and_hits and not all(and_hits)) or (or_hits and not any(or_hits)):
                        ok = False
                        break

                if ok:
                    ids.append(e.Id)

            if ids:
                lst = List[DB.ElementId](ids)
                t = Transaction(self.doc, "Isolate Filtered Elements")
                t.Start()
                self.curview.IsolateElementsTemporary(lst)
                t.Commit()
            else:
                dlg = TaskDialog("Warning")
                dlg.MainInstruction = "No elements match the selected filters."
                dlg.CommonButtons = TaskDialogCommonButtons.Ok
                dlg.Show()

        except Exception as e:
            dlg = TaskDialog("Error")
            dlg.MainInstruction = "Isolate Error: {0}".format(e)
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()

    def select_clicked(self, sender, args):
        if not self.selected_filters:
            dlg = TaskDialog("Warning")
            dlg.MainInstruction = "No filters selected to select."
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()
            return

        try:
            pre = [self.doc.GetElement(i) for i in self.uidoc.Selection.GetElementIds()]
            use_all = any(k in self.selected_filters for k in ("Name", "Comments", "Category", "TS_Point_Number", "TS_Point_Description"))
            elems = pre or (self.all_elements if use_all else self.fab_elements)

            ids = []
            for e in elems:
                if not e or not e.IsValidObject:
                    continue

                ev = {p: get_property_value(e, p, self.config, False) for p in self.selected_filters}
                ok = True

                for p, fl in self.selected_filters.items():
                    and_hits = []
                    or_hits = []
                    for vals, andf in fl:
                        hit = (str(ev[p]) in [str(v) for v in vals])
                        if andf:
                            and_hits.append(hit)
                        else:
                            or_hits.append(hit)

                    if (and_hits and not all(and_hits)) or (or_hits and not any(or_hits)):
                        ok = False
                        break

                if ok:
                    ids.append(e.Id)

            if ids:
                self.uidoc.Selection.SetElementIds(List[DB.ElementId](ids))
                self.Close()
            else:
                dlg = TaskDialog("Warning")
                dlg.MainInstruction = "No elements match the selected filters."
                dlg.CommonButtons = TaskDialogCommonButtons.Ok
                dlg.Show()

        except Exception as e:
            dlg = TaskDialog("Error")
            dlg.MainInstruction = "Select Error: {0}".format(e)
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()

    def cancel_clicked(self, sender, args):
        self.Close()


def get_property_value(elem, property_name, config, debug=False):
    if elem is None or not elem.IsValidObject:
        return None

    property_map = {
        'CID': lambda x: str(x.ItemCustomId) if x.ItemCustomId else None,
        'ServiceType': lambda x: config.GetServiceTypeName(x.ServiceType) if x.ServiceType else None,
        'Name': lambda x: get_parameter_value_by_name_AsValueString(x, 'Family') or
                          (x.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()
                           if x.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM) else None),
        'Service Name': lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'),
        'Service Abbreviation': lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'),
        'Size': lambda x: get_parameter_value_by_name_AsString(x, 'Size of Primary End'),
        'STRATUS Assembly': lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Assembly'),
        'Line Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Line Number'),
        'STRATUS Status': lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Status'),
        'Reference Level': lambda x: get_parameter_value_by_name_AsValueString(x, 'Reference Level'),
        'Item Number': lambda x: get_parameter_value_by_name_AsString(x, 'Item Number'),
        'Bundle Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Bundle'),
        'REF BS Designation': lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF BS Designation'),
        'REF Line Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF Line Number'),
        'Comments': lambda x: get_parameter_value_by_name_AsString(x, 'Comments'),
        'Specification': lambda x: config.GetSpecificationName(x.Specification) if x.Specification else None,
        'Hanger Rod Size': lambda x: get_parameter_value_by_name_AsValueString(x, 'FP_Rod Size'),
        'Valve Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Valve Number'),
        'Beam Hanger': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Beam Hanger'),
        'Product Entry': lambda x: get_parameter_value_by_name_AsString(x, 'Product Entry'),
        'TS_Point_Number': lambda x: get_parameter_value_by_name_AsString(x, 'TS_Point_Number'),
        'TS_Point_Description': lambda x: get_parameter_value_by_name_AsString(x, 'TS_Point_Description'),
        'Alias': lambda x: get_parameter_value_by_name_AsString(x, 'Alias'),
        'Category': lambda x: x.Category.Name if x.Category else None,
    }

    try:
        return property_map.get(property_name, lambda x: None)(elem)
    except:
        return None


def get_parameter_id(property_name):
    param_map = {
        'STRATUS Assembly': 'STRATUS Assembly',
        'Line Number': 'FP_Line Number',
        'Service Name': 'Fabrication Service Name',
        'Service Abbreviation': 'Fabrication Service Abbreviation',
        'Size': 'Size of Primary End',
        'STRATUS Status': 'STRATUS Status',
        'Reference Level': 'Reference Level',
        'Item Number': 'Item Number',
        'Bundle Number': 'FP_Bundle',
        'REF BS Designation': 'FP_REF BS Designation',
        'REF Line Number': 'FP_REF Line Number',
        'Comments': 'Comments',
        'Hanger Rod Size': 'FP_Rod Size',
        'Valve Number': 'FP_Valve Number',
        'Beam Hanger': 'FP_Beam Hanger',
        'Product Entry': 'Product Entry',
        'Name': 'Family',
        'TS_Point_Number': 'TS_Point_Number',
        'TS_Point_Description': 'TS_Point_Description',
        'Alias': 'Alias',
    }
    return param_map.get(property_name)


def run(uiapp):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        dlg = TaskDialog("Error")
        dlg.MainInstruction = "No active Revit document."
        dlg.CommonButtons = TaskDialogCommonButtons.Ok
        dlg.Show()
        return

    doc = uidoc.Document
    curview = doc.ActiveView
    config = FabricationConfiguration.GetFabricationConfiguration(doc)

    Shared_Params()

    preselection = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    fab_elements = preselection or FilteredElementCollector(
        doc, curview.Id
    ).OfClass(DB.FabricationPart).WhereElementIsNotElementType().ToElements()

    all_elements = preselection or FilteredElementCollector(
        doc, curview.Id
    ).WhereElementIsNotElementType().ToElements()

    property_options = {}

    fab_props = [
        'CID', 'ServiceType', 'Service Name', 'Service Abbreviation', 'Size',
        'STRATUS Assembly', 'Line Number', 'STRATUS Status', 'Reference Level',
        'Item Number', 'Bundle Number', 'REF BS Designation', 'REF Line Number',
        'Specification', 'Hanger Rod Size', 'Valve Number', 'Beam Hanger',
        'Product Entry', 'TS_Point_Number', 'TS_Point_Description', 'Alias'
    ]

    for prop in fab_props:
        vals = set(filter(None, [get_property_value(e, prop, config, False) for e in fab_elements]))
        if vals:
            property_options[prop] = sorted(vals)

    all_props = ['Name', 'Comments', 'Category', 'TS_Point_Number', 'TS_Point_Description']
    for prop in all_props:
        if all_elements:
            vals = set(filter(None, [get_property_value(e, prop, config, False) for e in all_elements]))
            if vals:
                property_options[prop] = sorted(vals)

    if not property_options:
        dlg = TaskDialog("Error")
        dlg.MainInstruction = "No properties found for the selected elements."
        dlg.CommonButtons = TaskDialogCommonButtons.Ok
        dlg.Show()
        return

    form = MultiPropertyFilterForm(doc, uidoc, curview, config, property_options, fab_elements, all_elements)
    form.frame = DispatcherFrame()
    form.Closed += form.exit_frame
    form.Show()
    Dispatcher.PushFrame(form.frame)