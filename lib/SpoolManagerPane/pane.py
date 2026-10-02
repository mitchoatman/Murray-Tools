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

from pyrevit import forms, HOST_APP
import Autodesk.Revit.UI as UI
from Autodesk.Revit.UI import UIThemeManager, UITheme

from System.Collections.Generic import List
from System.Windows.Data import CollectionViewSource
from System.Windows.Media import SolidColorBrush, Color as MediaColor
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    Transaction,
    TransactionGroup,
    FilteredElementCollector,
    ElementId,
    FabricationPart,
    ParameterFilterRuleFactory,
    Color,
    BuiltInCategory,
    ParameterFilterElement,
    ElementParameterFilter,
    OverrideGraphicSettings,
    View,
    FabricationConfiguration,
    IndependentTag,
    TagOrientation,
    Family,
    FamilySymbol,
    XYZ
)
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, FontWeights
from System.Windows.Controls import Button, TextBox, Label, Grid, RowDefinition, ColumnDefinition, ListBox, StackPanel, Orientation
from System.Windows.Media import FontFamily
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Forms import MessageBox

try:
    from Parameters.Add_SharedParameters import Shared_Params
except Exception:
    def Shared_Params():
        pass

try:
    from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString
except Exception:
    def set_parameter_by_name(elem, param_name, value):
        param = elem.LookupParameter(param_name)
        if param and not param.IsReadOnly:
            param.Set(value)
            
    def get_parameter_value_by_name_AsString(elem, param_name):
        param = elem.LookupParameter(param_name)
        if param:
            return param.AsString() or param.AsValueString() or ""
        return ""


PANE_XAML = """
<Page
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:scm="clr-namespace:System.ComponentModel;assembly=WindowsBase"
    x:Name="root_page"
    Background="{DynamicResource PageBackgroundBrush}">

    <Page.Resources>
        <!-- Dynamic Theme Brushes Setup (Defaults to Light) -->
        <SolidColorBrush x:Key="PageBackgroundBrush" Color="#FFF5F5F5"/>
        <SolidColorBrush x:Key="TextForegroundBrush" Color="#FF333333"/>
        <SolidColorBrush x:Key="SubTextForegroundBrush" Color="#FF666666"/>
        <SolidColorBrush x:Key="ControlBackgroundBrush" Color="#FFFFFFFF"/>
        <SolidColorBrush x:Key="ButtonBackgroundBrush" Color="#FFEFEFEF"/>
        <SolidColorBrush x:Key="BorderColorBrush" Color="#FFD0D0D0"/>
        <SolidColorBrush x:Key="SeparatorColorBrush" Color="#FF000000"/>
    </Page.Resources>

    <Grid Margin="10">
        <Grid.Resources>
            <Style TargetType="TextBlock">
                <Setter Property="Foreground" Value="{DynamicResource TextForegroundBrush}"/>
            </Style>
            <Style TargetType="TextBox">
                <Setter Property="Background" Value="{DynamicResource ControlBackgroundBrush}"/>
                <Setter Property="Foreground" Value="{DynamicResource TextForegroundBrush}"/>
                <Setter Property="BorderBrush" Value="{DynamicResource BorderColorBrush}"/>
                <Setter Property="Padding" Value="2,1,2,1"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="TextBox">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="1">
                                <ScrollViewer x:Name="PART_ContentHost"/>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style TargetType="Button">
                <Setter Property="Background" Value="{DynamicResource ButtonBackgroundBrush}"/>
                <Setter Property="Foreground" Value="{DynamicResource TextForegroundBrush}"/>
                <Setter Property="BorderBrush" Value="{DynamicResource BorderColorBrush}"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="1">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </Grid.Resources>

        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Two-Column Top Layout Matching User Reference -->
        <Grid Grid.Row="0" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Left Column: Set Spool Data Workflow -->
            <StackPanel Grid.Column="0">
                <TextBlock Text="Spool Name:" Margin="0,0,0,2"/>
                <TextBox x:Name="spool_input_tb" Height="22" Margin="0,0,0,6"/>
                <TextBlock Text="Map Name:" Margin="0,0,0,2"/>
                <TextBox x:Name="map_input_tb" Height="22" Margin="0,0,0,4"/>
                <Button x:Name="apply_btn" Content="Set Spool Data" Height="24"/>
            </StackPanel>

            <!-- Right Column: Auto Spool Workflow -->
            <StackPanel Grid.Column="2">
                <TextBlock Text="Spool Length (ft):" Margin="0,0,0,2"/>
                <TextBox x:Name="length_input_tb" Height="22" Margin="0,0,0,6"/>
                <TextBlock Text="Spool Width (ft):" Margin="0,0,0,2"/>
                <TextBox x:Name="width_input_tb" Height="22" Margin="0,0,0,4"/>
                <Button x:Name="auto_spool_btn" Content="Auto Spool" Height="24"/>
            </StackPanel>
        </Grid>

        <Separator Grid.Row="1"
                   Margin="0,4,0,8"
                   Background="{DynamicResource SeparatorColorBrush}"
                   Height="1"/>

        <!-- Package Assignment Section Mirroring Two-Column Widths -->
        <Grid Grid.Row="2" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Column="0">
                <TextBlock Text="Package Name:" Margin="0,0,0,2"/>
                <TextBox x:Name="package_input_tb" Height="22"/>
            </StackPanel>

            <StackPanel Grid.Column="2">
                <TextBlock Text="" Margin="0,0,0,2" Height="14"/>
                <Button x:Name="assign_package_btn" Content="Assign Package" Height="24"/>
            </StackPanel>
        </Grid>

        <Separator Grid.Row="3"
                   Margin="0,4,0,8"
                   Background="{DynamicResource SeparatorColorBrush}"
                   Height="1"/>

        <StackPanel Grid.Row="4" Margin="0,0,0,6">
            <TextBlock Text="Search List:"
                       Margin="0,0,0,2"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="24"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="search_tb"
                         Grid.Column="0"
                         Height="22"/>
                <Button x:Name="clear_search_btn"
                        Grid.Column="1"
                        Content="X"
                        Height="22"
                        Margin="4,0,0,0"/>
            </Grid>
        </StackPanel>

        <TextBlock Grid.Row="5"
                   Text="Packages &amp; Spools in View (Ctrl/Shift for Multi-Select, Double Click to Zoom)"
                   FontWeight="SemiBold"
                   Margin="0,0,0,4"/>

        <!-- Grouped ListBox with Nested Package Headers -->
        <ListBox x:Name="spool_list_box"
                 Grid.Row="7"
                 SelectionMode="Extended"
                 DisplayMemberPath="spool_name"
                 MinHeight="200"
                 BorderBrush="{DynamicResource BorderColorBrush}"
                 BorderThickness="1"
                 Background="{DynamicResource ControlBackgroundBrush}"
                 Foreground="{DynamicResource TextForegroundBrush}">
            <ListBox.GroupStyle>
                <GroupStyle>
                    <GroupStyle.HeaderTemplate>
                        <DataTemplate>
                            <TextBlock Text="{Binding Name}" FontWeight="Bold" Margin="2,4,2,2" Foreground="{DynamicResource TextForegroundBrush}"/>
                        </DataTemplate>
                    </GroupStyle.HeaderTemplate>
                </GroupStyle>
            </ListBox.GroupStyle>
        </ListBox>

        <!-- Balanced 50/50 Side-by-Side Buttons -->
        <Grid Grid.Row="8" Margin="0,4,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="4"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="select_all_btn"
                    Grid.Column="0"
                    Content="Select all in List"
                    Height="22"/>
            <Button x:Name="pin_all_btn"
                    Grid.Column="2"
                    Content="Pin Spools in View"
                    Height="22"/>
        </Grid>

        <Separator Grid.Row="9"
                   Margin="0,8,0,8"
                   Background="{DynamicResource SeparatorColorBrush}"
                   Height="1"/>

        <StackPanel Grid.Row="10" Margin="0,0,0,6">
            <TextBlock Text="Find / Replace (From selected spools in list):"
                       Margin="0,0,0,2"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="24"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="find_tb"
                         Grid.Column="0"
                         Height="22"
                         ToolTip="Find text"/>
                <Button x:Name="clear_find_btn"
                        Grid.Column="1"
                        Content="X"
                        Height="22"
                        Margin="4,0,0,0"/>
            </Grid>
            <TextBox x:Name="replace_tb"
                     Height="22"
                     Margin="0,4,0,0"
                     ToolTip="Replace text"/>
            <Button x:Name="rename_btn"
                    Content="Rename Selected"
                    Height="24"
                    Margin="0,4,0,0"/>
        </StackPanel>

        <Separator Grid.Row="11"
                   Margin="0,8,0,8"
                   Background="{DynamicResource SeparatorColorBrush}"
                   Height="1"/>

        <!-- Bottom Row Buttons Side-by-Side -->
        <Grid Grid.Row="12" Margin="0,0,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="4"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="make_filters_btn"
                    Grid.Column="0"
                    Content="Create Spool View Filters"
                    Height="26"/>
            <Button x:Name="tag_spools_btn"
                    Grid.Column="2"
                    Content="Tag Spools in View"
                    Height="26"/>
        </Grid>

        <TextBlock x:Name="status_tb"
                   Grid.Row="13"
                   Margin="0,6,0,0"
                   Foreground="{DynamicResource SubTextForegroundBrush}"
                   TextWrapping="Wrap"
                   Text="Ready."/>
    </Grid>
</Page>
"""


PARAM_NAME = "STRATUS Assembly"
PACKAGE_PARAM_NAME = "STRATUS Package"
MAP_PARAM_NAME = "FP_Spool Map"
FOLDER_NAME = r"C:\Temp"
FILE_PATH = os.path.join(FOLDER_NAME, "Ribbon_StratusAssembly.txt")
AUTOSPOOL_FILE_PATH = os.path.join(FOLDER_NAME, "Ribbon_StratusAssembly_AutoSpool.txt")

DEFAULT_SPOOL = "L1-A1-CW-01"
DEFAULT_MAP = "L1-A1-HGR-MAP"
DEFAULT_LENGTH = "20"
DEFAULT_WIDTH = "10"

state = UI.DockablePaneState()
state.DockPosition = UI.DockPosition.Right


class SpoolDisplayItem(System.Object):
    def __init__(self, package_name, spool_name):
        self._package_name = package_name
        self._spool_name = spool_name

    @property
    def package_name(self):
        return self._package_name

    @property
    def spool_name(self):
        return self._spool_name


def natural_key(value):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r"([0-9]+)", value or "")]


def ensure_input_file():
    if not os.path.exists(FOLDER_NAME):
        os.makedirs(FOLDER_NAME)

    if not os.path.exists(FILE_PATH):
        with open(FILE_PATH, "w") as f:
            f.writelines([
                "L1-A1-CW-01\n",
                "L1-A1-HGR-MAP\n"
            ])
            
    if not os.path.exists(AUTOSPOOL_FILE_PATH):
        with open(AUTOSPOOL_FILE_PATH, "w") as f:
            f.writelines([DEFAULT_SPOOL + "\n", DEFAULT_MAP + "\n", DEFAULT_LENGTH + "\n", DEFAULT_WIDTH + "\n"])


def read_previous_input():
    ensure_input_file()
    try:
        with open(FILE_PATH, "r") as f:
            lines = [x.strip() for x in f.readlines()]
            while len(lines) < 2:
                lines.append("")
            return lines[0], lines[1]
    except Exception:
        return "", ""


def write_previous_input(spool_name, map_name):
    ensure_input_file()
    with open(FILE_PATH, "w") as f:
        f.writelines([
            spool_name + "\n",
            map_name + "\n"
        ])


def read_autospool_defaults():
    ensure_input_file()
    try:
        with open(AUTOSPOOL_FILE_PATH, 'r') as f:
            lines = [x.rstrip() for x in f.readlines()]
        while len(lines) < 4:
            lines.append("")
        return lines[0] or DEFAULT_SPOOL, lines[1] or DEFAULT_MAP, lines[2] or DEFAULT_LENGTH, lines[3] or DEFAULT_WIDTH
    except Exception:
        return DEFAULT_SPOOL, DEFAULT_MAP, DEFAULT_LENGTH, DEFAULT_WIDTH


def write_autospool_defaults(spool_name, map_name, length_str, width_str):
    ensure_input_file()
    with open(AUTOSPOOL_FILE_PATH, 'w') as f:
        f.writelines([spool_name + "\n", map_name + "\n", str(length_str) + "\n", str(width_str) + "\n"])


def increment_spool_name(value):
    if not value:
        return value
    
    parts = value.rsplit('-', 1)
    if len(parts) == 2 and parts[1].isdigit():
        width = len(parts[1])
        next_num = int(parts[1]) + 1
        return "{}-{}".format(parts[0], str(next_num).zfill(width))

    match = list(re.finditer(r'\d+', value))
    if not match:
        return value
    
    last_match = match[-1]
    num_str = last_match.group()
    start, end = last_match.span()
    
    width = len(num_str)
    next_num = int(num_str) + 1
    next_num_str = str(next_num).zfill(width)
    
    return value[:start] + next_num_str + value[end:]


def get_connectors(el):
    try:
        cm = el.ConnectorManager or el.MEPModel.ConnectorManager
        return list(cm.Connectors) if cm else []
    except:
        return []


def get_element_dimensions(el):
    try:
        bbox = el.get_BoundingBox(None)
        if bbox:
            min_pt = bbox.Min
            max_pt = bbox.Max
            dx = abs(max_pt.X - min_pt.X)
            dy = abs(max_pt.Y - min_pt.Y)
            dz = abs(max_pt.Z - min_pt.Z)
            dims = sorted([dx, dy, dz], reverse=True)
            return dims[0], dims[1], dims[2]
    except:
        pass
    
    try:
        val = el.CenterlineLength
        if val:
            return float(val), 0.0, 0.0
    except:
        pass
        
    return 0.0, 0.0, 0.0


def dot(v1, v2):
    return v1.X * v2.X + v1.Y * v2.Y + v1.Z * v2.Z


def find_connection_in_set(conn, run_id_set, exclude_el_id):
    try:
        for rc in conn.AllRefs:
            owner = rc.Owner
            if owner and owner.Id.IntegerValue != exclude_el_id.IntegerValue and owner.Id.IntegerValue in run_id_set:
                return owner, rc
    except:
        pass
    return None, None


class RunWalkError(Exception):
    pass


def walk_run(doc, run_id_set, start_id):
    order = []
    branch_stubs = []
    visited = set()
    current_id = start_id
    incoming_connector = None

    while True:
        el = doc.GetElement(current_id)
        visited.add(current_id.IntegerValue)
        connectors = get_connectors(el)

        outgoing_candidates = [c for c in connectors if c is not incoming_connector] if incoming_connector else list(connectors)

        viable = []
        for c in outgoing_candidates:
            other_el, other_conn = find_connection_in_set(c, run_id_set, current_id)
            if other_el is None or other_el.Id.IntegerValue in visited:
                continue
            score = dot(incoming_connector.CoordinateSystem.BasisZ, c.CoordinateSystem.BasisZ) if incoming_connector else None
            viable.append((c, other_el, other_conn, score))

        length_dim, width_dim, _ = get_element_dimensions(el)
        order.append({'element': el, 'length': length_dim, 'width': width_dim})

        if not viable:
            break

        chosen = None
        if incoming_connector is None:
            if len(viable) == 1:
                chosen = viable[0]
            else:
                raise RunWalkError("Starting element has multiple directions. Pick an element at the physical end of the run.")
        elif len(connectors) >= 3:
            scored = [v for v in viable if v[3] is not None]
            scored.sort(key=lambda v: v[3])
            chosen = scored[0] if scored else viable[0]
            for v in scored[1:]:
                branch_stubs.append("Tee/cross at element {} -> unassigned branch at {}".format(current_id.IntegerValue, v[1].Id.IntegerValue))
        else:
            chosen = viable[0]

        if chosen is None:
            break

        current_id = chosen[1].Id
        incoming_connector = chosen[2]

    return order, branch_stubs


def group_into_spools(order, start_spool_name, target_length, target_width):
    groups, current_group, current_len, current_width, spool_name = [], [], 0.0, 0.0, start_spool_name
    for item in order:
        current_group.append(item['element'])
        current_len += item['length']
        if item['width'] > current_width:
            current_width = item['width']
            
        if current_len >= target_length or current_width >= target_width:
            groups.append((spool_name, current_group))
            spool_name = increment_spool_name(spool_name)
            current_group, current_len, current_width = [], 0.0, 0.0
            
    if current_group:
        groups.append((spool_name, current_group))
    return groups


def assign_spool_params(elements, spool_name, map_name, missing_elements):
    for el in elements:
        try:
            if el.LookupParameter("Fabrication Service") or el.LookupParameter("STRATUS Assembly"):
                set_parameter_by_name(el, "STRATUS Assembly", spool_name)
                set_parameter_by_name(el, "FP_Spool Map", map_name)
                set_parameter_by_name(el, "STRATUS Status", "Modeled")
                try: el.SpoolName = spool_name
                except: pass
                try: el.PartStatus = 1
                except: pass
                try: el.Pinned = True
                except: pass
            else:
                missing_elements.append(str(el.Id.IntegerValue))
        except Exception as ex:
            missing_elements.append("{} : {}".format(el.Id.IntegerValue, str(ex)))


class RunSelectionFilter(ISelectionFilter):
    def __init__(self, allowed_ids):
        self.allowed_ids = allowed_ids
    def AllowElement(self, elem):
        return elem.Id.IntegerValue in self.allowed_ids
    def AllowReference(self, ref, pos):
        return True


def get_spools_in_view(doc, curview):
    package_dict = {}
    if not curview:
        return package_dict

    categories = [
        BuiltInCategory.OST_FabricationPipework,
        BuiltInCategory.OST_FabricationDuctwork,
        BuiltInCategory.OST_FabricationHangers,
        BuiltInCategory.OST_FabricationContainment
    ]
    for cat in categories:
        collector = FilteredElementCollector(doc, curview.Id).OfCategory(cat).OfClass(FabricationPart)
        for elem in collector:
            try:
                param = elem.LookupParameter(PARAM_NAME)
                if param and param.HasValue:
                    spool_val = param.AsString()
                    if spool_val and spool_val.strip():
                        spool_val = spool_val.strip()
                        
                        pkg_param = elem.LookupParameter(PACKAGE_PARAM_NAME)
                        pkg_val = "Unassigned Package"
                        if pkg_param and pkg_param.HasValue:
                            p_str = pkg_param.AsString()
                            if p_str and p_str.strip():
                                pkg_val = p_str.strip()
                                
                        if pkg_val not in package_dict:
                            package_dict[pkg_val] = set()
                        package_dict[pkg_val].add(spool_val)
            except Exception:
                pass

    sorted_package_dict = {}
    for pkg in sorted(package_dict.keys(), key=natural_key):
        sorted_package_dict[pkg] = sorted(list(package_dict[pkg]), key=natural_key)

    return sorted_package_dict


def get_elements_by_spool(doc, curview, spool_name):
    matches = []
    if not curview:
        return matches

    categories = [
        BuiltInCategory.OST_FabricationPipework,
        BuiltInCategory.OST_FabricationDuctwork,
        BuiltInCategory.OST_FabricationHangers,
        BuiltInCategory.OST_FabricationContainment
    ]
    for cat in categories:
        collector = FilteredElementCollector(doc, curview.Id).OfCategory(cat).OfClass(FabricationPart)
        for elem in collector:
            try:
                param = elem.LookupParameter(PARAM_NAME)
                if param and param.HasValue and param.AsString() == spool_name:
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


def get_target_view_or_template(doc):
    curview = doc.ActiveView
    view_template_id = curview.ViewTemplateId
    if not view_template_id.Equals(ElementId.InvalidElementId):
        return doc.GetElement(view_template_id)
    return curview


class FamilyLoaderOptionsHandler(System.Object, System.IDisposable, DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

    def Dispose(self):
        pass


class _SpoolManagerRequestHandler(IExternalEventHandler):
    def __init__(self):
        self.request_name = None
        self.spool_name = None
        self.package_name = None
        self.map_name = None
        self.find_text = None
        self.replace_text = None
        self.selected_items = None
        self.pane = None

    def Execute(self, uiapp):
        if self.pane is None:
            return

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            self.pane.set_status("No active document.")
            self.pane.set_all_spools({})
            return

        doc = uidoc.Document
        curview = doc.ActiveView

        try:
            if self.request_name == "refresh":
                self._do_refresh(doc, curview)

            elif self.request_name == "show":
                self._do_show(uidoc, doc, curview, self.spool_name)

            elif self.request_name == "apply":
                self._do_apply(uidoc, doc, curview, self.spool_name, self.map_name)

            elif self.request_name == "assign_package":
                self._do_assign_package(doc, curview, self.package_name, self.selected_items)

            elif self.request_name == "auto_spool":
                self._do_auto_spool(uiapp, uidoc, doc, curview)

            elif self.request_name == "rename":
                self._do_rename(doc, curview, self.find_text, self.replace_text, self.selected_items)

            elif self.request_name == "make_filters":
                self._do_make_filters(doc, curview)

            elif self.request_name == "tag_spools":
                self._do_tag_spools(doc, curview)

            elif self.request_name == "pin_all":
                self._do_pin_all(doc, curview)

        except Exception as ex:
            self.pane.set_status("Error: {}".format(str(ex)))

        finally:
            self.request_name = None
            self.spool_name = None
            self.package_name = None
            self.map_name = None
            self.find_text = None
            self.replace_text = None
            self.selected_items = None

    def GetName(self):
        return "Spool Manager Pane External Event"

    def _do_refresh(self, doc, curview):
        packages_map = get_spools_in_view(doc, curview)
        total_spools = sum(len(spools) for spools in packages_map.values())
        self.pane.set_all_spools(packages_map)
        self.pane.set_status("{} spool(s) found across {} package(s) in active view.".format(total_spools, len(packages_map)))

    def _do_show(self, uidoc, doc, curview, spool_name):
        if not spool_name:
            self.pane.set_status("Select a spool name first.")
            return

        matches = get_elements_by_spool(doc, curview, spool_name)
        if not matches:
            self.pane.set_status("No elements found for spool '{}' in active view.".format(spool_name))
            return

        show_elements(uidoc, matches)
        self.pane.set_status("Showing {} element(s) for spool '{}'.".format(len(matches), spool_name))

    def _do_apply(self, uidoc, doc, curview, spool_name, map_name):
        spool_name = (spool_name or "").strip()
        map_name = (map_name or "").strip()
        
        if not spool_name or not map_name:
            self.pane.set_status("Enter both Spool Name and Map Name.")
            return

        selection_ids = list(uidoc.Selection.GetElementIds())
        if not selection_ids:
            self.pane.set_status("Select one or more elements in Revit, then click Set Spool Data.")
            return

        Shared_Params()
        FabricationConfiguration.GetFabricationConfiguration(doc)

        missing_elements = []
        changed = 0
        
        t = Transaction(doc, "Set Spool Data")
        t.Start()
        try:
            for eid in selection_ids:
                el = doc.GetElement(eid)
                if el is None:
                    continue

                try:
                    param_exist = el.LookupParameter(PARAM_NAME)
                    isfabpart = el.LookupParameter("Fabrication Service")

                    if isfabpart:
                        set_parameter_by_name(el, PARAM_NAME, spool_name)
                        set_parameter_by_name(el, MAP_PARAM_NAME, map_name)
                        set_parameter_by_name(el, "STRATUS Status", "Modeled")

                        try:
                            el.SpoolName = spool_name
                        except:
                            pass

                        try:
                            el.PartStatus = 1
                        except:
                            pass

                        try:
                            el.Pinned = True
                        except:
                            pass
                        changed += 1

                    elif param_exist:
                        set_parameter_by_name(el, PARAM_NAME, spool_name)
                        set_parameter_by_name(el, MAP_PARAM_NAME, map_name)
                        set_parameter_by_name(el, "STRATUS Status", "Modeled")

                        try:
                            el.Pinned = True
                        except:
                            pass
                        changed += 1
                    else:
                        missing_elements.append(str(eid.IntegerValue))
                except Exception as inner_ex:
                    missing_elements.append("{} : {}".format(eid.IntegerValue, str(inner_ex)))

            t.Commit()
        except Exception:
            if t.HasStarted():
                t.RollBack()
            raise

        new_spool_name = increment_spool_name(spool_name)
        write_previous_input(new_spool_name, map_name)
        
        self.pane.spool_input_tb.Text = new_spool_name
        self.pane.map_input_tb.Text = map_name

        packages_map = get_spools_in_view(doc, curview)
        self.pane.set_all_spools(packages_map)
        self.pane.set_status("Applied spool data to {} element(s). Next spool incremented.".format(changed))

    def _do_assign_package(self, doc, curview, package_name, selected_spools):
        package_name = (package_name or "").strip()
        if not package_name:
            self.pane.set_status("Enter a Package Name first.")
            return
        if not selected_spools:
            self.pane.set_status("Select one or more spools in the list.")
            return

        categories = [
            BuiltInCategory.OST_FabricationPipework,
            BuiltInCategory.OST_FabricationDuctwork,
            BuiltInCategory.OST_FabricationHangers,
            BuiltInCategory.OST_FabricationContainment
        ]

        elements_to_process = []
        for cat in categories:
            collector = FilteredElementCollector(doc, curview.Id).OfCategory(cat).OfClass(FabricationPart)
            for elem in collector:
                try:
                    asm_val = get_parameter_value_by_name_AsString(elem, PARAM_NAME)
                    if asm_val and asm_val in selected_spools:
                        elements_to_process.append(elem)
                except Exception:
                    pass

        if not elements_to_process:
            self.pane.set_status("No fabrication parts matched the selected spools.")
            return

        changed = 0
        t = Transaction(doc, "Assign STRATUS Package")
        t.Start()
        try:
            for elem in elements_to_process:
                set_parameter_by_name(elem, PACKAGE_PARAM_NAME, package_name)
                changed += 1
            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            self.pane.set_status("Assign package error: {}".format(str(ex)))
            return

        packages_map = get_spools_in_view(doc, curview)
        self.pane.set_all_spools(packages_map)
        self.pane.set_status("Assigned package '{}' to {} element(s).".format(package_name, changed))

    def _do_auto_spool(self, uiapp, uidoc, doc, curview):
        run_ids = list(uidoc.Selection.GetElementIds())
        if not run_ids:
            self.pane.set_status("No elements selected. Multi-select the pipe run first.")
            return

        run_id_set = set(eid.IntegerValue for eid in run_ids)

        spool_name = self.pane.spool_input_tb.Text.strip()
        map_name = self.pane.map_input_tb.Text.strip()
        length_text = self.pane.length_input_tb.Text.strip()
        width_text = self.pane.width_input_tb.Text.strip()

        if not spool_name or not map_name or not length_text or not width_text:
            self.pane.set_status("Enter Spool Name, Map Name, Length, and Width.")
            return

        try:
            target_length = float(length_text)
            target_width = float(width_text)
            if target_length <= 0 or target_width <= 0:
                raise ValueError()
        except ValueError:
            self.pane.set_status("Length and Width must be positive numbers.")
            return

        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, RunSelectionFilter(run_id_set), "Pick starting element at physical end of run.")
        except:
            self.pane.set_status("Auto spool selection cancelled.")
            return

        order, branch_stubs = walk_run(doc, run_id_set, ref.ElementId)

        if not order:
            self.pane.set_status("Could not walk run.")
            return

        groups = group_into_spools(order, spool_name, target_length, target_width)

        missing_elements = []
        t = Transaction(doc, "Auto Spool Run")
        t.Start()
        for gname, gelements in groups:
            assign_spool_params(gelements, gname, map_name, missing_elements)
        t.Commit()

        next_spool = increment_spool_name(groups[-1][0]) if groups else spool_name
        write_autospool_defaults(next_spool, map_name, length_text, width_text)
        write_previous_input(next_spool, map_name)

        self.pane.spool_input_tb.Text = next_spool
        self.pane.map_input_tb.Text = map_name

        packages_map = get_spools_in_view(doc, curview)
        self.pane.set_all_spools(packages_map)
        self.pane.set_status("Auto spool complete: Assigned {} spool(s).".format(len(groups)))

    def _do_rename(self, doc, curview, find, replace, selected_items):
        if not find or replace is None:
            self.pane.set_status("Please enter both Find and Replace values.")
            return
        if not selected_items:
            self.pane.set_status("No spools selected in list.")
            return

        categories = [
            BuiltInCategory.OST_FabricationPipework,
            BuiltInCategory.OST_FabricationDuctwork,
            BuiltInCategory.OST_FabricationHangers,
            BuiltInCategory.OST_FabricationContainment
        ]

        elements_to_process = []
        for cat in categories:
            collector = FilteredElementCollector(doc, curview.Id).OfCategory(cat).OfClass(FabricationPart)
            for elem in collector:
                try:
                    asm_val = get_parameter_value_by_name_AsString(elem, PARAM_NAME)
                    if asm_val and asm_val in selected_items:
                        elements_to_process.append(elem)
                except Exception:
                    pass

        if not elements_to_process:
            self.pane.set_status("No fabrication parts matched your selection.")
            return

        changed = 0
        t = Transaction(doc, "Rename STRATUS Assemblies")
        t.Start()
        try:
            for elem in elements_to_process:
                current = get_parameter_value_by_name_AsString(elem, PARAM_NAME)
                param = elem.LookupParameter(PARAM_NAME)
                if current and param and not param.IsReadOnly:
                    new_val = current.replace(find, replace)
                    set_parameter_by_name(elem, PARAM_NAME, new_val)
                    try:
                        elem.SpoolName = new_val
                    except Exception:
                        pass
                    changed += 1
            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            self.pane.set_status("Rename error: {}".format(str(ex)))
            return

        packages_map = get_spools_in_view(doc, curview)
        self.pane.set_all_spools(packages_map)
        self.pane.set_status("Renamed spools on {} element(s).".format(changed))

    def _do_make_filters(self, doc, curview):
            Shared_Params()
            app = doc.Application
            RevitINT = float(app.VersionNumber)
            view_to_modify = get_target_view_or_template(doc)

            part_collector = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                                        .WhereElementIsNotElementType() \
                                        .ToElements()

            existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
            existing_filter_names = {filter.Name for filter in existing_filters}
            existing_filter_dict = {filter.Name: filter.Id for filter in existing_filters}

            from collections import OrderedDict
            custom_filters = OrderedDict()
            custom_filters["SPOOL 1"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "1",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (255, 0, 0)
            }
            custom_filters["SPOOL 2"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "2",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (79, 0, 59)
            }
            custom_filters["SPOOL 3"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "3",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (0, 0, 255)
            }
            custom_filters["SPOOL 4"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "4",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (0, 255, 0)
            }
            custom_filters["SPOOL 5"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "5",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (255, 0, 0)
            }
            custom_filters["SPOOL 6"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "6",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (79, 0, 59)
            }
            custom_filters["SPOOL 7"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "7",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (0, 0, 255)
            }
            custom_filters["SPOOL 8"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "8",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (0, 255, 0)
            }
            custom_filters["SPOOL 9"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "9",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (255, 0, 0)
            }
            custom_filters["SPOOL 0"] = {
                "parameter_name": "STRATUS Assembly", "condition": "EndsWith", "value": "0",
                "categories": [BuiltInCategory.OST_FabricationPipework, BuiltInCategory.OST_FabricationHangers, BuiltInCategory.OST_FabricationDuctwork],
                "color": (79, 0, 59)
            }

            applied_filter_ids = view_to_modify.GetFilters()

            with Transaction(doc, "Create and Apply Spool Filters") as t:
                t.Start()
                for filter_name, filter_props in custom_filters.items():
                    param_name = filter_props["parameter_name"]
                    condition = filter_props["condition"]
                    value = filter_props["value"]
                    cats = List[ElementId]([ElementId(cat) for cat in filter_props["categories"]])
                    color = Color(*filter_props["color"])

                    if filter_name in existing_filter_names:
                        filter_id = existing_filter_dict[filter_name]
                    else:
                        sample_element = part_collector[0] if part_collector else None
                        if not sample_element:
                            continue

                        param_id = None
                        for p in sample_element.Parameters:
                            if p.Definition.Name == param_name:
                                param_id = p.Id
                                break
                        if not param_id:
                            continue

                        if RevitINT < 2023:
                            rule = ParameterFilterRuleFactory.CreateEndsWithRule(param_id, value, False)
                        else:
                            rule = ParameterFilterRuleFactory.CreateEndsWithRule(param_id, value)

                        filter_element = ElementParameterFilter(rule)
                        filter_elem = ParameterFilterElement.Create(doc, filter_name, cats)
                        filter_elem.SetElementFilter(filter_element)
                        filter_id = filter_elem.Id
                        existing_filter_dict[filter_name] = filter_id

                    if not applied_filter_ids.Contains(filter_id):
                        view_to_modify.AddFilter(filter_id)

                    view_to_modify.SetFilterVisibility(filter_id, True)
                    overrides = OverrideGraphicSettings()
                    overrides.SetProjectionLineColor(color)
                    view_to_modify.SetFilterOverrides(filter_id, overrides)

                t.Commit()

            self.pane.set_status("Spool view filters created and missing ones applied successfully.")

    def _do_tag_spools(self, doc, curview):
        Shared_Params()
        app = doc.Application
        RevitINT = float(app.VersionNumber)

        script_dir = os.path.dirname(__file__)
        family_name = 'Fabrication Pipe - Stratus Assembly'
        family_filename = 'Fabrication Pipe - Stratus Assembly.rfa'
        family_path = os.path.join(script_dir, family_filename)
        
        families = FilteredElementCollector(doc).OfClass(Family)
        fam_is_in_project = any(f.Name == family_name for f in families)
        
        if not fam_is_in_project:
            if os.path.exists(family_path):
                with Transaction(doc, 'Load Stratus Family Tag') as t:
                    t.Start()
                    fload_handler = FamilyLoaderOptionsHandler()
                    doc.LoadFamily(family_path, fload_handler)
                    t.Commit()
            else:
                self.pane.set_status("Family file not found at: {}".format(family_path))
                return

        all_tags = FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements()
        tag_type = next((tag for tag in all_tags if tag.Family.Name == family_name), None)
        if not tag_type:
            self.pane.set_status("Tag family '{}' not found in project.".format(family_name))
            return

        part_collector = FilteredElementCollector(doc, curview.Id)\
            .OfClass(FabricationPart)\
            .WhereElementIsNotElementType()\
            .ToElements()

        if not part_collector:
            self.pane.set_status("No Fabrication parts found in the active view.")
            return

        def get_safe_stratus(element):
            try:
                val = get_parameter_value_by_name_AsString(element, 'STRATUS Assembly')
                return val if val else "Unknown"
            except:
                return "Unknown"

        existing_tags = FilteredElementCollector(doc, curview.Id).OfClass(IndependentTag).ToElements()
        tagged_spools = set()
        for tag in existing_tags:
            try:
                if RevitINT > 2021:
                    for element_id in tag.GetTaggedLocalElementIds():
                        tagged_elem = doc.GetElement(element_id)
                        if tagged_elem and isinstance(tagged_elem, FabricationPart):
                            spool_val = get_safe_stratus(tagged_elem)
                            if spool_val != "Unknown":
                                tagged_spools.add(spool_val)
                else:
                    element_id = tag.TaggedLocalElementId
                    tagged_elem = doc.GetElement(element_id)
                    if tagged_elem and isinstance(tagged_elem, FabricationPart):
                        spool_val = get_safe_stratus(tagged_elem)
                        if spool_val != "Unknown":
                            tagged_spools.add(spool_val)
            except:
                pass

        elementlist = []
        tagged_assemblies = set()

        for elem in part_collector:
            assembly_value = get_safe_stratus(elem)
            if (assembly_value != "Unknown" and 
                "-MAP" not in assembly_value and
                assembly_value not in tagged_assemblies and 
                assembly_value not in tagged_spools and
                elem.LookupParameter("Fabrication Service")):
                elementlist.append(elem)
                tagged_assemblies.add(assembly_value)

        if not elementlist:
            self.pane.set_status("No untagged elements found for tagging.")
            return

        tagged_count = 0
        with TransactionGroup(doc, "Tag Spools") as tg:
            tg.Start()
            for element in elementlist:
                try:
                    location = element.Location
                    tag_point = None

                    if isinstance(location, DB.LocationPoint):
                        tag_point = location.Point
                    elif isinstance(location, DB.LocationCurve) and location.Curve:
                        tag_point = location.Curve.Evaluate(0.5, True)

                    if not tag_point:
                        for next_elem in part_collector:
                            if next_elem != element:
                                next_location = next_elem.Location
                                if isinstance(next_location, DB.LocationCurve) and next_location.Curve:
                                    tag_point = next_location.Curve.Evaluate(0.9, True)
                                    if tag_point:
                                        break

                    if tag_point:
                        with Transaction(doc, "Tag Fabrication Part") as t:
                            t.Start()
                            IndependentTag.Create(
                                doc,
                                tag_type.Id,
                                curview.Id,
                                DB.Reference(element),
                                True,
                                TagOrientation.Horizontal,
                                tag_point
                            )
                            t.Commit()
                        tagged_count += 1
                except:
                    pass
            tg.Assimilate()

        self.pane.set_status("Successfully tagged {} spool(s).".format(tagged_count))

    def _do_pin_all(self, doc, curview):
        all_elements = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                                   .WhereElementIsNotElementType() \
                                   .ToElements()

        pinned_count = 0
        with Transaction(doc, 'Pin All Spools in View') as t:
            t.Start()
            for i in all_elements:
                param_exist = i.LookupParameter("STRATUS Assembly")
                if param_exist and param_exist.HasValue:
                    try:
                        i.Pinned = True
                        pinned_count += 1
                    except:
                        pass
            t.Commit()

        self.pane.set_status("Pinned {} fabrication part(s) in active view.".format(pinned_count))


class SpoolManagerPane(forms.WPFPanel):
    panel_id = "8D4E9D3F-8B1E-4B7E-8C4D-7F5B2F0F62B3"
    panel_title = "Spool Manager"
    panel_source = "inline"
    initial_state = state

    def __init__(self):
        self.load_xaml(PANE_XAML, literal_string=True)

        self._handler = _SpoolManagerRequestHandler()
        self._handler.pane = self
        self._ext_event = ExternalEvent.Create(self._handler)

        self._all_packages_map = {}

        spool_def, map_def = read_previous_input()
        auto_spool_def, auto_map_def, auto_len_def, auto_width_def = read_autospool_defaults()
        
        self.spool_input_tb.Text = spool_def
        self.map_input_tb.Text = map_def
        self.length_input_tb.Text = auto_len_def
        self.width_input_tb.Text = auto_width_def

        self.apply_btn.Click += self.on_apply_clicked
        self.assign_package_btn.Click += self.on_assign_package_clicked
        self.auto_spool_btn.Click += self.on_auto_spool_clicked
        self.rename_btn.Click += self.on_rename_clicked
        self.make_filters_btn.Click += self.on_make_filters_clicked
        self.tag_spools_btn.Click += self.on_tag_spools_clicked
        self.pin_all_btn.Click += self.on_pin_all_clicked
        
        self.clear_find_btn.Click += lambda s, e: setattr(self.find_tb, 'Text', '')
        self.clear_search_btn.Click += lambda s, e: setattr(self.search_tb, 'Text', '')
        self.select_all_btn.Click += self.on_select_all_clicked

        self.search_tb.TextChanged += self.on_search_changed
        self.spool_list_box.SelectionChanged += self.on_list_item_selected
        self.spool_list_box.MouseDoubleClick += self.on_list_double_clicked
        
        self.Loaded += self.on_loaded
        self.Unloaded += self.on_unloaded

        try:
            from pyrevit import HOST_APP
            HOST_APP.uiapp.ViewActivated += self.on_view_activated
            app = HOST_APP.app
            app.DocumentChanged += self.on_document_changed

            # Check Revit version for 2024+ native dark theme support
            if HOST_APP.is_newer_than(2023, True):
                HOST_APP.uiapp.ThemeChanged += self.on_theme_changed
                self.apply_revit_theme()
            else:
                # Force strictly to Light Mode palette for Revit 2023 and older
                self.apply_light_theme()
        except Exception:
            self.apply_light_theme()

    def apply_light_theme(self):
        try:
            res = self.Resources
            res["PageBackgroundBrush"] = SolidColorBrush(MediaColor.FromArgb(255, 245, 245, 245))
            res["TextForegroundBrush"] = SolidColorBrush(MediaColor.FromArgb(255, 51, 51, 51))
            res["SubTextForegroundBrush"] = SolidColorBrush(MediaColor.FromArgb(255, 102, 102, 102))
            res["ControlBackgroundBrush"] = SolidColorBrush(MediaColor.FromArgb(255, 255, 255, 255))
            res["ButtonBackgroundBrush"] = SolidColorBrush(MediaColor.FromArgb(255, 239, 239, 239))
            res["BorderColorBrush"] = SolidColorBrush(MediaColor.FromArgb(255, 208, 208, 208))
            res["SeparatorColorBrush"] = SolidColorBrush(MediaColor.FromArgb(255, 0, 0, 0))
        except Exception:
            pass

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
            sep_hex = "#FFFFFFFF" if is_dark else "#FF000000"

            res = self.Resources
            res["PageBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["TextForegroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(text_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["SubTextForegroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(subtext_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["ControlBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(ctrl_bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["ButtonBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(btn_bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["BorderColorBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(border_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["SeparatorColorBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(sep_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
        except Exception:
            pass

    def on_theme_changed(self, sender, args):
        self.apply_revit_theme()

    def on_loaded(self, sender, args):
        self.request_refresh()

    def on_unloaded(self, sender, args):
        try:
            from pyrevit import HOST_APP
            HOST_APP.uiapp.ViewActivated -= self.on_view_activated
            app = HOST_APP.app
            app.DocumentChanged -= self.on_document_changed
            if HOST_APP.is_newer_than(2023, True):
                HOST_APP.uiapp.ThemeChanged -= self.on_theme_changed
        except Exception:
            pass

    def on_view_activated(self, sender, args):
        self.request_refresh()

    def on_document_changed(self, sender, args):
        self.request_refresh()

    def on_list_item_selected(self, sender, args):
        selected = self.spool_list_box.SelectedItem
        if selected and isinstance(selected, SpoolDisplayItem):
            self.spool_input_tb.Text = selected.spool_name
            if selected.package_name and selected.package_name != "Unassigned Package":
                self.package_input_tb.Text = selected.package_name

    def on_list_double_clicked(self, sender, args):
        selected = self.spool_list_box.SelectedItem
        if selected and isinstance(selected, SpoolDisplayItem):
            self.request_show(selected.spool_name)

    def on_apply_clicked(self, sender, args):
        self.request_apply(self.spool_input_tb.Text, self.map_input_tb.Text)

    def on_assign_package_clicked(self, sender, args):
        selected_spools = []
        for item in self.spool_list_box.SelectedItems:
            if isinstance(item, SpoolDisplayItem):
                selected_spools.append(item.spool_name)
        self.request_assign_package(self.package_input_tb.Text, selected_spools)

    def on_auto_spool_clicked(self, sender, args):
        self.request_auto_spool()

    def on_rename_clicked(self, sender, args):
        selected_spools = []
        for item in self.spool_list_box.SelectedItems:
            if isinstance(item, SpoolDisplayItem):
                selected_spools.append(item.spool_name)
        self.request_rename(self.find_tb.Text, self.replace_tb.Text, selected_spools)

    def on_select_all_clicked(self, sender, args):
        self.spool_list_box.SelectAll()

    def on_pin_all_clicked(self, sender, args):
        self.request_pin_all()

    def on_make_filters_clicked(self, sender, args):
        self.request_make_filters()

    def on_tag_spools_clicked(self, sender, args):
        self.request_tag_spools()

    def on_search_changed(self, sender, args):
        self.apply_search_filter()

    def request_refresh(self):
        self._handler.request_name = "refresh"
        self._handler.spool_name = None
        self._handler.package_name = None
        self._handler.map_name = None
        try:
            self._ext_event.Raise()
        except Exception:
            pass

    def request_show(self, spool_name):
        self._handler.request_name = "show"
        self._handler.spool_name = spool_name
        self._handler.package_name = None
        self._handler.map_name = None
        self._ext_event.Raise()

    def request_apply(self, spool_name, map_name):
        self._handler.request_name = "apply"
        self._handler.spool_name = spool_name
        self._handler.map_name = map_name
        self._ext_event.Raise()

    def request_assign_package(self, package_name, selected_items):
        self._handler.request_name = "assign_package"
        self._handler.package_name = package_name
        self._handler.selected_items = selected_items
        self._ext_event.Raise()

    def request_auto_spool(self):
        self._handler.request_name = "auto_spool"
        self._handler.spool_name = None
        self._handler.package_name = None
        self._handler.map_name = None
        self._ext_event.Raise()

    def request_rename(self, find_text, replace_text, selected_items):
        self._handler.request_name = "rename"
        self._handler.find_text = find_text
        self._handler.replace_text = replace_text
        self._handler.selected_items = selected_items
        self._ext_event.Raise()

    def request_make_filters(self):
        self._handler.request_name = "make_filters"
        self._handler.spool_name = None
        self._handler.package_name = None
        self._handler.map_name = None
        self._ext_event.Raise()

    def request_tag_spools(self):
        self._handler.request_name = "tag_spools"
        self._handler.spool_name = None
        self._handler.package_name = None
        self._handler.map_name = None
        self._ext_event.Raise()

    def request_pin_all(self):
        self._handler.request_name = "pin_all"
        self._handler.spool_name = None
        self._handler.package_name = None
        self._handler.map_name = None
        self._ext_event.Raise()

    def set_all_spools(self, packages_map):
        self._all_packages_map = packages_map
        self.apply_search_filter()

    def apply_search_filter(self):
        current_text = self.spool_input_tb.Text
        search_text = (self.search_tb.Text or "").strip().lower()

        items_list = []
        for package_name, spools in sorted(self._all_packages_map.items(), key=lambda x: natural_key(x[0])):
            for spool in spools:
                if not search_text or search_text in spool.lower() or search_text in package_name.lower():
                    items_list.append(SpoolDisplayItem(package_name, spool))

        cvs = CollectionViewSource()
        cvs.Source = items_list
        cvs.GroupDescriptions.Add(System.Windows.Data.PropertyGroupDescription("package_name"))
        self.spool_list_box.ItemsSource = cvs.View

        self.spool_input_tb.Text = current_text

    def set_status(self, message):
        self.status_tb.Text = message or ""