# -*- coding: utf-8 -*-
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Collections')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System import TimeSpan, Action
from System.Windows import (
    Window, WindowStartupLocation, WindowStyle, GridLength,
    HorizontalAlignment, VerticalAlignment, GridUnitType, Thickness,
    Visibility, TextWrapping
)
from System.Windows.Controls import (
    Label, ComboBox, Button, ListBox, CheckBox, TextBox, Grid,
    RowDefinition, ColumnDefinition, SelectionMode, StackPanel,
    Orientation, ListBoxItem, Border, TextBlock
)
from System.Windows.Input import Cursors
from System.Windows.Media import Brushes, SolidColorBrush, Color
from System.Collections.Generic import List
from System.Windows.Threading import (
    DispatcherFrame, Dispatcher, DispatcherTimer, DispatcherPriority
)

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


FAB_ONLY_SCAN_PROPS = [
    'CID', 'ServiceType', 'Service Name', 'Service Abbreviation', 'Size',
    'STRATUS Assembly', 'Line Number', 'STRATUS Status', 'Reference Level',
    'Item Number', 'Bundle Number', 'REF BS Designation', 'REF Line Number',
    'Specification', 'Insulation Specification', 'Hanger Rod Size',
    'Valve Number', 'Beam Hanger', 'Product Entry', 'Alias', 'Cut Type',
    'Part Material', 'Pointload'
]

ALL_SCAN_PROPS = [
    'Name', 'Comments', 'Category', 'TS_Point_Number', 'TS_Point_Description'
]


def has_value(v):
    return v not in (None, "")


def safe_str(v):
    return "" if v is None else str(v)


_MISSING = object()

_ELEMENT_TYPE_CACHE = {}
_INSTANCE_PARAM_STRING_CACHE = {}
_TYPE_PARAM_STRING_CACHE = {}
_INSTANCE_PARAM_VALUESTRING_CACHE = {}
_TYPE_PARAM_VALUESTRING_CACHE = {}
_INSTANCE_PARAM_INT_CACHE = {}
_TYPE_PARAM_INT_CACHE = {}

_PROPERTY_VALUE_CACHE = {}
_NAME_VALUE_CACHE = {}
_WORKSET_NAME_CACHE = {}
_SERVICE_TYPE_NAME_CACHE = {}
_SPEC_NAME_CACHE = {}
_INSUL_SPEC_CACHE = {}


def clear_global_caches():
    _ELEMENT_TYPE_CACHE.clear()
    _INSTANCE_PARAM_STRING_CACHE.clear()
    _TYPE_PARAM_STRING_CACHE.clear()
    _INSTANCE_PARAM_VALUESTRING_CACHE.clear()
    _TYPE_PARAM_VALUESTRING_CACHE.clear()
    _INSTANCE_PARAM_INT_CACHE.clear()
    _TYPE_PARAM_INT_CACHE.clear()
    _PROPERTY_VALUE_CACHE.clear()
    _NAME_VALUE_CACHE.clear()
    _WORKSET_NAME_CACHE.clear()
    _SERVICE_TYPE_NAME_CACHE.clear()
    _SPEC_NAME_CACHE.clear()
    _INSUL_SPEC_CACHE.clear()


def get_elem_key(elem):
    try:
        if elem and elem.IsValidObject:
            return elem.Id.IntegerValue
    except:
        pass
    return None


def get_cached_param_string(target, param_name, cache):
    if target is None or not target.IsValidObject:
        return None

    try:
        key = (target.Id.IntegerValue, param_name)
    except:
        return None

    if key in cache:
        v = cache[key]
        return None if v is _MISSING else v

    try:
        v = get_parameter_value_by_name_AsString(target, param_name)
        cache[key] = v if has_value(v) else _MISSING
        return v if has_value(v) else None
    except:
        cache[key] = _MISSING
        return None


def get_cached_param_valuestring(target, param_name, cache):
    if target is None or not target.IsValidObject:
        return None

    try:
        key = (target.Id.IntegerValue, param_name)
    except:
        return None

    if key in cache:
        v = cache[key]
        return None if v is _MISSING else v

    try:
        v = get_parameter_value_by_name_AsValueString(target, param_name)
        cache[key] = v if has_value(v) else _MISSING
        return v if has_value(v) else None
    except:
        cache[key] = _MISSING
        return None


def get_cached_param_int(target, param_name, cache):
    if target is None or not target.IsValidObject:
        return None

    try:
        key = (target.Id.IntegerValue, param_name)
    except:
        return None

    if key in cache:
        v = cache[key]
        return None if v is _MISSING else v

    try:
        v = get_parameter_value_by_name_AsInteger(target, param_name)
        cache[key] = v if v is not None else _MISSING
        return v
    except:
        cache[key] = _MISSING
        return None


def get_cached_service_type_name(config, service_type_id):
    key = safe_str(service_type_id)
    if key in _SERVICE_TYPE_NAME_CACHE:
        v = _SERVICE_TYPE_NAME_CACHE[key]
        return None if v is _MISSING else v

    try:
        v = config.GetServiceTypeName(service_type_id)
    except:
        v = None

    _SERVICE_TYPE_NAME_CACHE[key] = v if has_value(v) else _MISSING
    return v


def get_cached_spec_name(config, spec_id):
    key = safe_str(spec_id)
    if key in _SPEC_NAME_CACHE:
        v = _SPEC_NAME_CACHE[key]
        return None if v is _MISSING else v

    try:
        v = config.GetSpecificationName(spec_id)
    except:
        v = None

    _SPEC_NAME_CACHE[key] = v if has_value(v) else _MISSING
    return v


def get_cached_insul_spec_abbrev(config, insul_spec_id):
    key = safe_str(insul_spec_id)
    if key in _INSUL_SPEC_CACHE:
        v = _INSUL_SPEC_CACHE[key]
        return None if v is _MISSING else v

    try:
        v = config.GetInsulationSpecificationAbbreviation(insul_spec_id)
    except:
        v = None

    _INSUL_SPEC_CACHE[key] = v if has_value(v) else _MISSING
    return v


def get_element_type(elem):
    if elem is None or not elem.IsValidObject:
        return None

    key = get_elem_key(elem)
    if key is not None and key in _ELEMENT_TYPE_CACHE:
        v = _ELEMENT_TYPE_CACHE[key]
        return None if v is _MISSING else v

    try:
        type_id = elem.GetTypeId()
        if type_id and type_id != DB.ElementId.InvalidElementId:
            t = elem.Document.GetElement(type_id)
            if key is not None:
                _ELEMENT_TYPE_CACHE[key] = t if t else _MISSING
            return t
    except:
        pass

    if key is not None:
        _ELEMENT_TYPE_CACHE[key] = _MISSING
    return None


def get_param_string_instance_or_type(elem, param_name):
    if elem is None or not elem.IsValidObject:
        return None

    v = get_cached_param_string(elem, param_name, _INSTANCE_PARAM_STRING_CACHE)
    if has_value(v):
        return v

    elem_type = get_element_type(elem)
    if elem_type:
        return get_cached_param_string(elem_type, param_name, _TYPE_PARAM_STRING_CACHE)

    return None


def get_param_value_string_instance_or_type(elem, param_name):
    if elem is None or not elem.IsValidObject:
        return None

    v = get_cached_param_valuestring(elem, param_name, _INSTANCE_PARAM_VALUESTRING_CACHE)
    if has_value(v):
        return v

    elem_type = get_element_type(elem)
    if elem_type:
        return get_cached_param_valuestring(elem_type, param_name, _TYPE_PARAM_VALUESTRING_CACHE)

    return None


def get_param_int_instance_or_type(elem, param_name):
    if elem is None or not elem.IsValidObject:
        return None

    v = get_cached_param_int(elem, param_name, _INSTANCE_PARAM_INT_CACHE)
    if v is not None:
        return v

    elem_type = get_element_type(elem)
    if elem_type:
        return get_cached_param_int(elem_type, param_name, _TYPE_PARAM_INT_CACHE)

    return None


def get_name_value(x):
    key = get_elem_key(x)
    if key is not None and key in _NAME_VALUE_CACHE:
        v = _NAME_VALUE_CACHE[key]
        return None if v is _MISSING else v

    val = get_param_value_string_instance_or_type(x, 'Family')
    if not val:
        try:
            p = x.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
            if p:
                val = p.AsValueString()
        except:
            pass

    if not val:
        elem_type = get_element_type(x)
        if elem_type:
            try:
                p = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if p:
                    val = p.AsString() or p.AsValueString()
            except:
                pass

    if key is not None:
        _NAME_VALUE_CACHE[key] = val if has_value(val) else _MISSING

    return val


def get_user_workset_name(elem):
    if elem is None or not elem.IsValidObject:
        return None

    key = get_elem_key(elem)
    if key is not None and key in _WORKSET_NAME_CACHE:
        v = _WORKSET_NAME_CACHE[key]
        return None if v is _MISSING else v

    val = None
    try:
        ws = elem.Document.GetWorksetTable().GetWorkset(elem.WorksetId)
        if ws and ws.Kind == DB.WorksetKind.UserWorkset:
            val = ws.Name
    except:
        pass

    if key is not None:
        _WORKSET_NAME_CACHE[key] = val if has_value(val) else _MISSING

    return val


PROPERTY_MAP = {
    'CID': lambda x, c: str(x.ItemCustomId) if getattr(x, 'ItemCustomId', None) else None,
    'ServiceType': lambda x, c: get_cached_service_type_name(c, x.ServiceType) if getattr(x, 'ServiceType', None) else None,
    'Name': lambda x, c: get_name_value(x),
    'Service Name': lambda x, c: get_param_string_instance_or_type(x, 'Fabrication Service Name'),
    'Service Abbreviation': lambda x, c: get_param_string_instance_or_type(x, 'Fabrication Service Abbreviation'),
    'Size': lambda x, c: get_param_string_instance_or_type(x, 'Size of Primary End'),
    'STRATUS Assembly': lambda x, c: get_param_string_instance_or_type(x, 'STRATUS Assembly'),
    'Line Number': lambda x, c: get_param_string_instance_or_type(x, 'FP_Line Number'),
    'STRATUS Status': lambda x, c: get_param_string_instance_or_type(x, 'STRATUS Status'),
    'Reference Level': lambda x, c: get_param_value_string_instance_or_type(x, 'Reference Level'),
    'Item Number': lambda x, c: get_param_string_instance_or_type(x, 'Item Number'),
    'Bundle Number': lambda x, c: get_param_string_instance_or_type(x, 'FP_Bundle'),
    'REF BS Designation': lambda x, c: get_param_string_instance_or_type(x, 'FP_REF BS Designation'),
    'REF Line Number': lambda x, c: get_param_string_instance_or_type(x, 'FP_REF Line Number'),
    'Comments': lambda x, c: get_param_string_instance_or_type(x, 'Comments'),
    'Specification': lambda x, c: get_cached_spec_name(c, x.Specification) if getattr(x, 'Specification', None) else None,
    'Part Material': lambda x, c: c.GetMaterialName(x.Material) if getattr(x, 'Material', None) is not None and c else None,
    'Hanger Rod Size': lambda x, c: get_param_value_string_instance_or_type(x, 'FP_Rod Size'),
    'Valve Number': lambda x, c: get_param_string_instance_or_type(x, 'FP_Valve Number'),
    'Beam Hanger': lambda x, c: get_param_string_instance_or_type(x, 'FP_Beam Hanger'),
    'Product Entry': lambda x, c: get_param_string_instance_or_type(x, 'Product Entry'),
    'TS_Point_Number': lambda x, c: get_param_string_instance_or_type(x, 'TS_Point_Number'),
    'TS_Point_Description': lambda x, c: get_param_string_instance_or_type(x, 'TS_Point_Description'),
    'Alias': lambda x, c: get_param_string_instance_or_type(x, 'Alias'),
    'Category': lambda x, c: x.Category.Name if x.Category else None,
    'Workset': lambda x, c: get_user_workset_name(x),
    'Cut Type': lambda x, c: get_param_value_string_instance_or_type(x, 'Cut Type'),
    'Insulation Specification': lambda x, c:
        get_param_value_string_instance_or_type(x, 'Insulation Specification') or
        (get_cached_insul_spec_abbrev(c, x.InsulationSpecification)
         if getattr(x, 'InsulationSpecification', 0) else None),
    'Pointload': lambda x, c: get_param_value_string_instance_or_type(x, 'FP_Pointload') or get_param_string_instance_or_type(x, 'FP_Pointload'),
}


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
    def __init__(self, doc, uidoc, curview, config, property_names, fab_elements, all_elements):
        self.doc = doc
        self.uidoc = uidoc
        self.curview = curview
        self.config = config
        self.fab_elements = fab_elements
        self.all_elements = all_elements
        self.selected_filters = {}

        self.property_names = sorted(property_names)
        self.property_options = dict((p, None) for p in self.property_names)

        self._property_string_cache = {}
        self._valid_workset_cache = {}
        self._search_loaded_all = False
        self._search_delay_ms = 300

        self.search_timer = DispatcherTimer()
        self.search_timer.Interval = TimeSpan.FromMilliseconds(self._search_delay_ms)
        self.search_timer.Tick += self.on_search_timer_tick

        self.notice_timer = DispatcherTimer()
        self.notice_timer.Interval = TimeSpan.FromMilliseconds(2200)
        self.notice_timer.Tick += self.on_notice_timer_tick

        self.InitializeComponents()

        if self.property_names:
            self.property_combo.SelectedItem = self.property_names[0]
            self.update_values_list(None, None)

        self.update_filter_display()

    def InitializeComponents(self):
        self.Title = "Multi-Property Filter"
        self.Width = 550
        self.Height = 625
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
        for prop in self.property_names:
            self.property_combo.Items.Add(prop)
        self.property_combo.SelectionChanged += self.on_property_changed
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
        self.search_box.TextChanged += self.on_search_text_changed
        Grid.SetRow(self.search_box, 1)
        Grid.SetColumn(self.search_box, 1)
        grid.Children.Add(self.search_box)
        self.search_box.Focus()

        self.values_label = Label(
            Content="Select Values:",
            Margin=Thickness(10, 5, 10, 0),
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
            Margin=Thickness(10, 5, 10, 0),
            HorizontalAlignment=HorizontalAlignment.Right
        )
        self.add_button.Click += self.add_filter
        Grid.SetRow(self.add_button, 2)
        Grid.SetColumn(self.add_button, 1)
        grid.Children.Add(self.add_button)

        self.logic_check = CheckBox(
            Content="AND logic (unchecked = OR)",
            Width=230,
            Margin=Thickness(10, 5, 10, 0),
            HorizontalAlignment=HorizontalAlignment.Center
        )
        Grid.SetRow(self.logic_check, 4)
        Grid.SetColumn(self.logic_check, 1)
        grid.Children.Add(self.logic_check)

        self.filter_button = Button(
            Width=500,
            Height=130,
            Margin=Thickness(10, 10, 10, 0),
            HorizontalContentAlignment=HorizontalAlignment.Stretch,
            VerticalContentAlignment=VerticalAlignment.Stretch,
            HorizontalAlignment=HorizontalAlignment.Center
        )
        self.filter_button.ToolTip = "Click here to modify Filters"
        self.filter_button.Click += self.remove_filter

        self.filter_text = TextBlock()
        self.filter_text.TextWrapping = TextWrapping.Wrap
        self.filter_text.Margin = Thickness(6, 4, 6, 4)
        self.filter_text.Width = 470

        self.filter_button.Content = self.filter_text

        Grid.SetRow(self.filter_button, 5)
        Grid.SetColumn(self.filter_button, 1)
        grid.Children.Add(self.filter_button)

        panel = StackPanel(
            Orientation=Orientation.Horizontal,
            HorizontalAlignment=HorizontalAlignment.Center,
            Margin=Thickness(0, 15, 0, 0)
        )
        Grid.SetRow(panel, 6)
        Grid.SetColumn(panel, 1)
        grid.Children.Add(panel)

        for txt, handler in [
            ("Reset View", self.reset_clicked),
            ("Isolate", self.isolate_clicked),
            ("Select", self.select_clicked),
            ("Close", self.cancel_clicked)
        ]:
            btn = Button(
                Content=txt,
                Width=90 if txt == "Reset View" else 80,
                Height=25,
                Margin=Thickness(5, 0, 5, 0)
            )
            if txt == "Reset View":
                btn.Background = SolidColorBrush(Color.FromRgb(160, 82, 82))
                btn.Foreground = Brushes.White
            btn.Click += handler
            panel.Children.Add(btn)

        self.loading_overlay = Grid()
        self.loading_overlay.Background = SolidColorBrush(Color.FromArgb(125, 0, 0, 0))
        self.loading_overlay.Visibility = Visibility.Collapsed
        self.loading_overlay.Margin = Thickness(10, 8, 10, 0)

        overlay_border = Border()
        overlay_border.Width = 320
        overlay_border.Height = 70
        overlay_border.Background = SolidColorBrush(Color.FromRgb(245, 245, 245))
        overlay_border.BorderBrush = SolidColorBrush(Color.FromRgb(160, 160, 160))
        overlay_border.BorderThickness = Thickness(1)
        overlay_border.HorizontalAlignment = HorizontalAlignment.Center
        overlay_border.VerticalAlignment = VerticalAlignment.Center

        self.loading_text = Label(
            Content="Loading data...",
            HorizontalAlignment=HorizontalAlignment.Center,
            HorizontalContentAlignment=HorizontalAlignment.Center,
            VerticalContentAlignment=VerticalAlignment.Center,
            Foreground=Brushes.Black,
            FontSize=14,
            Margin=Thickness(10, 10, 10, 10)
        )

        overlay_border.Child = self.loading_text
        self.loading_overlay.Children.Add(overlay_border)

        Grid.SetRow(self.loading_overlay, 4)
        Grid.SetRowSpan(self.loading_overlay, 3)
        Grid.SetColumnSpan(self.loading_overlay, 3)
        grid.Children.Add(self.loading_overlay)

        self.notice_overlay = Grid()
        self.notice_overlay.Visibility = Visibility.Collapsed
        self.notice_overlay.Margin = Thickness(10, 8, 10, 0)

        self.notice_border = Border()
        self.notice_border.Width = 360
        self.notice_border.Height = 70
        self.notice_border.Background = SolidColorBrush(Color.FromRgb(255, 248, 225))
        self.notice_border.BorderBrush = SolidColorBrush(Color.FromRgb(191, 144, 0))
        self.notice_border.BorderThickness = Thickness(1)
        self.notice_border.HorizontalAlignment = HorizontalAlignment.Center
        self.notice_border.VerticalAlignment = VerticalAlignment.Center

        self.notice_text = Label(
            Content="",
            HorizontalAlignment=HorizontalAlignment.Center,
            HorizontalContentAlignment=HorizontalAlignment.Center,
            VerticalContentAlignment=VerticalAlignment.Center,
            Foreground=Brushes.Black,
            FontSize=13,
            Margin=Thickness(10, 10, 10, 10)
        )

        self.notice_border.Child = self.notice_text
        self.notice_overlay.Children.Add(self.notice_border)

        Grid.SetRow(self.notice_overlay, 4)
        Grid.SetRowSpan(self.notice_overlay, 3)
        Grid.SetColumnSpan(self.notice_overlay, 3)
        grid.Children.Add(self.notice_overlay)

    def exit_frame(self, s, e):
        self.frame.Continue = False

    def pump_ui(self):
        self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))

    def show_loading(self, text="Loading data..."):
        self.hide_notice()
        self.loading_text.Content = text
        self.loading_overlay.Visibility = Visibility.Visible
        self.Cursor = Cursors.Wait
        self.pump_ui()

    def hide_loading(self):
        self.loading_overlay.Visibility = Visibility.Collapsed
        self.Cursor = Cursors.Arrow
        self.pump_ui()

    def show_notice(self, text, level="warning", auto_hide=True):
        self.notice_timer.Stop()
        self.notice_text.Content = text

        if level == "warning":
            self.notice_border.Background = SolidColorBrush(Color.FromRgb(255, 248, 225))
            self.notice_border.BorderBrush = SolidColorBrush(Color.FromRgb(191, 144, 0))
        elif level == "info":
            self.notice_border.Background = SolidColorBrush(Color.FromRgb(232, 244, 253))
            self.notice_border.BorderBrush = SolidColorBrush(Color.FromRgb(91, 155, 213))
        else:
            self.notice_border.Background = SolidColorBrush(Color.FromRgb(240, 240, 240))
            self.notice_border.BorderBrush = SolidColorBrush(Color.FromRgb(160, 160, 160))

        self.notice_overlay.Visibility = Visibility.Visible
        self.pump_ui()

        if auto_hide:
            self.notice_timer.Start()

    def hide_notice(self):
        self.notice_timer.Stop()
        self.notice_overlay.Visibility = Visibility.Collapsed

    def on_notice_timer_tick(self, sender, args):
        self.notice_timer.Stop()
        self.hide_notice()

    def set_busy(self, is_busy, text=None):
        self.values_list.IsEnabled = not is_busy
        self.property_combo.IsEnabled = not is_busy
        self.add_button.IsEnabled = not is_busy
        self.logic_check.IsEnabled = not is_busy
        self.filter_button.IsEnabled = not is_busy
        self.search_box.IsEnabled = not is_busy

        if is_busy:
            self.show_loading(text or "Loading data...")
        else:
            self.hide_loading()

    def clear_runtime_caches(self):
        self._property_string_cache = {}
        self._valid_workset_cache = {}
        self._search_loaded_all = False
        clear_global_caches()

    def refresh_dialog_data(self):
        pre = [self.doc.GetElement(i) for i in self.uidoc.Selection.GetElementIds()]

        self.fab_elements = pre or FilteredElementCollector(
            self.doc, self.curview.Id
        ).OfClass(DB.FabricationPart).WhereElementIsNotElementType().ToElements()

        self.all_elements = pre or FilteredElementCollector(
            self.doc, self.curview.Id
        ).WhereElementIsNotElementType().ToElements()

        self.property_names = build_relevant_property_names(
            self.fab_elements,
            self.all_elements,
            self.config
        )
        self.property_options = dict((p, None) for p in self.property_names)
        self.clear_runtime_caches()

        current_sel = self.property_combo.SelectedItem
        self.property_combo.Items.Clear()
        for p in self.property_names:
            self.property_combo.Items.Add(p)

        self.values_list.Items.Clear()

        if self.property_names:
            if current_sel in self.property_names:
                self.property_combo.SelectedItem = current_sel
            else:
                self.property_combo.SelectedItem = self.property_names[0]
            self.update_values_list(None, None)

    def on_search_text_changed(self, sender, args):
        term = self.search_box.Text.strip()
        self.search_timer.Stop()

        if term:
            self.show_loading("Preparing search...")
            self.search_timer.Start()
        else:
            self.hide_loading()
            self.update_values_list(None, None)

    def on_property_changed(self, sender, args):
        try:
            self.property_combo.IsDropDownOpen = False
        except:
            pass

        def do_update():
            if self.search_box.Text.strip():
                self.search_timer.Stop()
                self.show_loading("Preparing search...")
                self.search_timer.Start()
            else:
                self.update_values_list(None, None)

        self.Dispatcher.BeginInvoke(
            DispatcherPriority.Background,
            Action(do_update)
        )

    def on_search_timer_tick(self, sender, args):
        self.search_timer.Stop()
        self.update_values_list(None, None)

    def ensure_property_loaded(self, prop):
        if prop not in self.property_options:
            return

        if self.property_options[prop] is not None:
            return

        vals = set()

        if prop in FAB_ONLY_SCAN_PROPS:
            elems = self.fab_elements
        else:
            elems = self.all_elements

        for e in elems:
            if prop == 'Workset':
                if not self.is_cached_valid_workset_element(e):
                    continue

            v = get_property_value(e, prop, self.config, False)
            if has_value(v):
                vals.add(v)

        self.property_options[prop] = sorted(vals, key=lambda x: safe_str(x).lower())

    def update_values_list(self, sender, args):
        term = self.search_box.Text.lower().strip()
        sel = self.property_combo.SelectedItem
        self.values_list.Items.Clear()

        try:
            if sel:
                self.set_busy(True, "Loading {0} values...".format(sel))
                self.ensure_property_loaded(sel)

                vals = self.property_options.get(sel, [])

                if term:
                    vals = [v for v in vals if term in safe_str(v).lower()]

                for v in vals:
                    vi = ValueItem(sel, v)
                    item = ListBoxItem(Content=str(vi))
                    item.Tag = vi
                    self.values_list.Items.Add(item)

                if term and not vals:
                    self.show_notice("No matching values found in {0}.".format(sel), "info")

        finally:
            self.set_busy(False)

    def add_filter(self, sender, args):
        sels = self.values_list.SelectedItems
        if not sels:
            self.show_notice("Please select at least one value.", "warning")
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
            self.show_notice("No filters to remove.", "info")
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
            self.filter_text.Text = "No Filters Yet..."
        else:
            txt = "Filters (Click here to modify):\n"
            for prop, fl in self.selected_filters.items():
                txt += "{0}:\n".format(prop)
                conds = []
                for vals, andf in fl:
                    mode = "AND" if andf else "OR"
                    conds.append("[{0}: {1}]".format(mode, ", ".join(str(v) for v in vals)))
                txt += " ".join(conds) + "\n"
            self.filter_text.Text = txt.strip()

    def reset_clicked(self, sender, args):
        try:
            self.set_busy(True, "Resetting view...")
            t = Transaction(self.doc, "Reset Temporary Hide/Isolate")
            t.Start()
            self.curview.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
            t.Commit()
            self.refresh_dialog_data()
        except Exception as e:
            dlg = TaskDialog("Error")
            dlg.MainInstruction = "Reset Error: {0}".format(e)
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()
        finally:
            self.set_busy(False)

    def is_cached_valid_workset_element(self, elem):
        if elem is None or not elem.IsValidObject:
            return False

        try:
            key = elem.Id.IntegerValue
        except:
            return False

        if key in self._valid_workset_cache:
            return self._valid_workset_cache[key]

        val = is_valid_workset_element(elem)
        self._valid_workset_cache[key] = val
        return val

    def get_cached_property_string(self, elem, prop):
        if elem is None or not elem.IsValidObject:
            return ""

        try:
            key = elem.Id.IntegerValue
        except:
            return ""

        prop_cache = self._property_string_cache.get(prop)
        if prop_cache is None:
            prop_cache = {}
            self._property_string_cache[prop] = prop_cache

        if key in prop_cache:
            return prop_cache[key]

        val = safe_str(get_property_value(elem, prop, self.config, False))
        prop_cache[key] = val
        return val

    def get_filtered_element_ids(self):
        pre = [self.doc.GetElement(i) for i in self.uidoc.Selection.GetElementIds()]
        use_all = any(
            p in ("Name", "Comments", "Category", "Workset", "TS_Point_Number", "TS_Point_Description")
            for p in self.selected_filters
        )
        elems = pre or (self.all_elements if use_all else self.fab_elements)

        and_filters = []
        or_filters = []

        for prop, flist in self.selected_filters.items():
            for vals, andf in flist:
                entry = (prop, set(str(v) for v in vals))
                if andf:
                    and_filters.append(entry)
                else:
                    or_filters.append(entry)

        if any(p == "Workset" for p, _ in and_filters + or_filters):
            elems = [e for e in elems if self.is_cached_valid_workset_element(e)]

        needed_props = set(p for p, _ in and_filters + or_filters)
        ids = []

        for e in elems:
            if not e or not e.IsValidObject:
                continue

            prop_vals = {}
            for p in needed_props:
                prop_vals[p] = self.get_cached_property_string(e, p)

            # every AND filter must pass
            and_ok = all(prop_vals[p] in valset for p, valset in and_filters)

            # if any OR filters exist, at least one must pass
            or_ok = True if not or_filters else any(
                prop_vals[p] in valset for p, valset in or_filters
            )

            if and_ok and or_ok:
                ids.append(e.Id)

        return ids

    def isolate_clicked(self, sender, args):
        if not self.selected_filters:
            self.show_notice("No filters selected to isolate.", "warning")
            return

        try:
            self.set_busy(True, "Filtering elements...")
            ids = self.get_filtered_element_ids()

            if ids:
                lst = List[DB.ElementId](ids)
                t = Transaction(self.doc, "Isolate Filtered Elements")
                t.Start()
                self.curview.IsolateElementsTemporary(lst)
                t.Commit()

                self.set_busy(True, "Refreshing data...")
                self.refresh_dialog_data()
            else:
                self.show_notice("No elements match the selected filters.", "warning")

        except Exception as e:
            dlg = TaskDialog("Error")
            dlg.MainInstruction = "Isolate Error: {0}".format(e)
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()
        finally:
            self.set_busy(False)

    def select_clicked(self, sender, args):
        if not self.selected_filters:
            self.show_notice("No filters selected to select.", "warning")
            return

        try:
            self.set_busy(True, "Filtering elements...")
            ids = self.get_filtered_element_ids()

            if ids:
                self.uidoc.Selection.SetElementIds(List[DB.ElementId](ids))
                self.Close()
            else:
                self.show_notice("No elements match the selected filters.", "warning")

        except Exception as e:
            dlg = TaskDialog("Error")
            dlg.MainInstruction = "Select Error: {0}".format(e)
            dlg.CommonButtons = TaskDialogCommonButtons.Ok
            dlg.Show()
        finally:
            self.set_busy(False)

    def cancel_clicked(self, sender, args):
        self.search_timer.Stop()
        self.notice_timer.Stop()
        self.Close()


def get_property_value(elem, property_name, config, debug=False):
    if elem is None or not elem.IsValidObject:
        return None

    elem_key = get_elem_key(elem)
    cache_key = (elem_key, property_name)

    if elem_key is not None and cache_key in _PROPERTY_VALUE_CACHE:
        v = _PROPERTY_VALUE_CACHE[cache_key]
        return None if v is _MISSING else v

    fn = PROPERTY_MAP.get(property_name)
    if not fn:
        return None

    try:
        val = fn(elem, config)
    except:
        val = None

    if elem_key is not None:
        _PROPERTY_VALUE_CACHE[cache_key] = val if has_value(val) else _MISSING

    return val


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
        'Part Material': 'Material',
        'Hanger Rod Size': 'FP_Rod Size',
        'Valve Number': 'FP_Valve Number',
        'Beam Hanger': 'FP_Beam Hanger',
        'Product Entry': 'Product Entry',
        'Name': 'Family',
        'Workset': 'Workset',
        'TS_Point_Number': 'TS_Point_Number',
        'TS_Point_Description': 'TS_Point_Description',
        'Alias': 'Alias',
        'Insulation Specification': 'Insulation Specification',
        'FP_Pointload': 'FP_Pointload',
    }
    return param_map.get(property_name)


def geometry_has_3d_objects(geo_elem):
    if geo_elem is None:
        return False

    for g in geo_elem:
        if isinstance(g, DB.Solid):
            try:
                if g.Volume > 0 or g.Faces.Size > 0:
                    return True
            except:
                pass

        elif isinstance(g, DB.Mesh):
            try:
                if g.NumTriangles > 0:
                    return True
            except:
                pass

        elif isinstance(g, DB.GeometryInstance):
            try:
                inst_geo = g.GetInstanceGeometry()
                if geometry_has_3d_objects(inst_geo):
                    return True
            except:
                pass

    return False


def has_3d_geometry(elem):
    if elem is None or not elem.IsValidObject:
        return False

    try:
        opt = DB.Options()
        opt.IncludeNonVisibleObjects = False
        geo = elem.get_Geometry(opt)
        return geometry_has_3d_objects(geo)
    except:
        return False


def is_valid_workset_element(elem):
    if elem is None or not elem.IsValidObject:
        return False

    try:
        if isinstance(elem, DB.View):
            return False
    except:
        pass

    try:
        if elem.ViewSpecific:
            return False
    except:
        pass

    cat = elem.Category
    if cat is None:
        return False

    try:
        if cat.CategoryType != DB.CategoryType.Model:
            return False
    except:
        pass

    if isinstance(elem, DB.Level):
        return False
    if isinstance(elem, DB.Grid):
        return False
    if isinstance(elem, DB.ReferencePlane):
        return False

    if not get_user_workset_name(elem):
        return False

    if not has_3d_geometry(elem):
        return False

    return True


def is_basic_workset_element(elem):
    if elem is None or not elem.IsValidObject:
        return False

    try:
        if isinstance(elem, DB.View):
            return False
    except:
        pass

    try:
        if elem.ViewSpecific:
            return False
    except:
        pass

    cat = elem.Category
    if cat is None:
        return False

    try:
        if cat.CategoryType != DB.CategoryType.Model:
            return False
    except:
        pass

    return bool(get_user_workset_name(elem))


def build_relevant_property_names(fab_elements, all_elements, config):
    relevant = set()

    remaining_fab = set(FAB_ONLY_SCAN_PROPS)
    for e in fab_elements:
        if not remaining_fab:
            break

        found_this_elem = []
        for prop in remaining_fab:
            v = get_property_value(e, prop, config, False)
            if has_value(v):
                relevant.add(prop)
                found_this_elem.append(prop)

        for prop in found_this_elem:
            remaining_fab.remove(prop)

    remaining_all = set(ALL_SCAN_PROPS)
    workset_found = False

    for e in all_elements:
        if not remaining_all and workset_found:
            break

        found_this_elem = []
        for prop in remaining_all:
            v = get_property_value(e, prop, config, False)
            if has_value(v):
                relevant.add(prop)
                found_this_elem.append(prop)

        for prop in found_this_elem:
            remaining_all.remove(prop)

        if not workset_found and is_basic_workset_element(e):
            relevant.add('Workset')
            workset_found = True

    return sorted(relevant)


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
    clear_global_caches()

    preselection = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]

    fab_elements = preselection or FilteredElementCollector(
        doc, curview.Id
    ).OfClass(DB.FabricationPart).WhereElementIsNotElementType().ToElements()

    all_elements = preselection or FilteredElementCollector(
        doc, curview.Id
    ).WhereElementIsNotElementType().ToElements()

    property_names = sorted(set(FAB_ONLY_SCAN_PROPS + ALL_SCAN_PROPS + ['Workset']))

    form = MultiPropertyFilterForm(
        doc,
        uidoc,
        curview,
        config,
        property_names,
        fab_elements,
        all_elements
    )
    form.frame = DispatcherFrame()
    form.Closed += form.exit_frame
    form.Show()
    Dispatcher.PushFrame(form.frame)