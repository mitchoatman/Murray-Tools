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
    FamilySymbol
)
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

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
    Background="#FFF5F5F5">

    <Grid Margin="10">
        <Grid.RowDefinitions>
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

        <StackPanel Grid.Row="0" Margin="0,0,0,6">
            <TextBlock Text="Spool Name:"
                       Margin="0,0,0,2"/>
            <TextBox x:Name="spool_input_tb"
                     Height="22"/>
        </StackPanel>

        <StackPanel Grid.Row="1" Margin="0,0,0,6">
            <TextBlock Text="Map Name:"
                       Margin="0,0,0,2"/>
            <TextBox x:Name="map_input_tb"
                     Height="22"/>
            <Button x:Name="apply_btn"
                    Content="Set Spool Data"
                    Height="24"
                    Margin="0,4,0,0"/>
        </StackPanel>

        <Separator Grid.Row="2"
                   Margin="0,4,0,8"
                   Background="#FF000000"
                   Height="1"/>

        <StackPanel Grid.Row="3" Margin="0,0,0,6">
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

        <TextBlock Grid.Row="4"
                   Text="Spools in View (Double Click to Zoom)"
                   FontWeight="SemiBold"
                   Margin="0,0,0,4"/>

        <ScrollViewer Grid.Row="6"
                      VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled"
                      MinHeight="220"
                      BorderBrush="#FFD0D0D0"
                      BorderThickness="1"
                      Background="White">
            <ListBox x:Name="line_numbers_lb"
                     SelectionMode="Extended"
                     BorderThickness="0"/>
        </ScrollViewer>

        <!-- Select All (3/4) and Pin All (1/4) Side-by-Side -->
        <Grid Grid.Row="7" Margin="0,4,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="4"/>
                <ColumnDefinition Width="1*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="select_all_btn"
                    Grid.Column="0"
                    Content="Select All"
                    Height="22"/>
            <Button x:Name="pin_all_btn"
                    Grid.Column="2"
                    Content="Pin All"
                    Height="22"/>
        </Grid>

        <Separator Grid.Row="8"
                   Margin="0,8,0,8"
                   Background="#FF000000"
                   Height="1"/>

        <StackPanel Grid.Row="9" Margin="0,0,0,6">
            <TextBlock Text="Find / Replace (Rename):"
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

        <Separator Grid.Row="10"
                   Margin="0,8,0,8"
                   Background="#FF000000"
                   Height="1"/>

        <!-- Bottom Row Buttons Side-by-Side -->
        <Grid Grid.Row="11" Margin="0,0,0,0">
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
                   Grid.Row="12"
                   Margin="0,6,0,0"
                   Foreground="#666666"
                   TextWrapping="Wrap"
                   Text="Ready."/>
    </Grid>
</Page>
"""


PARAM_NAME = "STRATUS Assembly"
MAP_PARAM_NAME = "FP_Spool Map"
FOLDER_NAME = r"C:\Temp"
FILE_PATH = os.path.join(FOLDER_NAME, "Ribbon_StratusAssembly.txt")

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
            f.writelines([
                "L1-A1-CW-01\n",
                "L1-A1-HGR-MAP\n"
            ])


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


def increment_spool_name(value):
    if not value:
        return value
    
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


def get_spools_in_view(doc, curview):
    values = set()
    if not curview:
        return sorted(values, key=natural_key)

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
                    val = param.AsString()
                    if val and val.strip():
                        values.add(val.strip())
            except Exception:
                pass

    return sorted(values, key=natural_key)


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
            self.pane.set_all_spools([])
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

            elif self.request_name == "rename":
                self._do_rename(doc, curview, self.find_text, self.replace_text, self.selected_items)

            elif self.request_name == "make_filters":
                self._do_make_filters(doc, curview)

            elif self.request_name == "tag_spools":
                self._do_tag_spools(doc, curview)

            elif self.request_name == "pin_all":
                self._do_pin_all(doc)

        except Exception as ex:
            self.pane.set_status("Error: {}".format(str(ex)))

        finally:
            self.request_name = None
            self.spool_name = None
            self.map_name = None
            self.find_text = None
            self.replace_text = None
            self.selected_items = None

    def GetName(self):
        return "Spool Manager Pane External Event"

    def _do_refresh(self, doc, curview):
        values = get_spools_in_view(doc, curview)
        self.pane.set_all_spools(values)
        self.pane.set_status("{} spool(s) found in active view.".format(len(values)))

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

        values = get_spools_in_view(doc, curview)
        self.pane.set_all_spools(values)
        self.pane.set_status("Applied spool data to {} element(s). Next spool incremented.".format(changed))

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

        values = get_spools_in_view(doc, curview)
        self.pane.set_all_spools(values)
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

                    # 1. Create or get the filter element globally in the document
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

                    # 2. Add only if it is missing from the view, then update overrides/visibility
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

    def _do_pin_all(self, doc):
        all_elements = FilteredElementCollector(doc).OfClass(FabricationPart) \
                             .WhereElementIsNotElementType() \
                             .ToElements()

        pinned_count = 0
        with Transaction(doc, 'Pin All Spools') as t:
            t.Start()
            for i in all_elements:
                param_exist = i.LookupParameter("STRATUS Assembly")
                if param_exist and param_exist.HasValue:
                    i.Pinned = True
                    pinned_count += 1
            t.Commit()

        self.pane.set_status("Pinned {} fabrication part(s).".format(pinned_count))


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

        self._all_spools = []

        spool_def, map_def = read_previous_input()
        self.spool_input_tb.Text = spool_def
        self.map_input_tb.Text = map_def

        self.apply_btn.Click += self.on_apply_clicked
        self.rename_btn.Click += self.on_rename_clicked
        self.make_filters_btn.Click += self.on_make_filters_clicked
        self.tag_spools_btn.Click += self.on_tag_spools_clicked
        self.pin_all_btn.Click += self.on_pin_all_clicked
        
        self.clear_find_btn.Click += lambda s, e: setattr(self.find_tb, 'Text', '')
        self.clear_search_btn.Click += lambda s, e: setattr(self.search_tb, 'Text', '')
        self.select_all_btn.Click += self.on_select_all_clicked

        self.search_tb.TextChanged += self.on_search_changed
        self.line_numbers_lb.SelectionChanged += self.on_spool_selected
        self.line_numbers_lb.MouseDoubleClick += self.on_spool_double_clicked
        self.line_numbers_lb.PreviewMouseWheel += self.on_list_mouse_wheel
        
        self.Loaded += self.on_loaded
        self.Unloaded += self.on_unloaded

        try:
            from pyrevit import HOST_APP
            HOST_APP.uiapp.ViewActivated += self.on_view_activated
        except Exception:
            pass

    def on_loaded(self, sender, args):
        self.request_refresh()

    def on_unloaded(self, sender, args):
        try:
            from pyrevit import HOST_APP
            HOST_APP.uiapp.ViewActivated -= self.on_view_activated
        except Exception:
            pass

    def on_view_activated(self, sender, args):
        self.request_refresh()

    def on_spool_selected(self, sender, args):
        selected = self.line_numbers_lb.SelectedItem
        if selected:
            self.spool_input_tb.Text = str(selected)

    def on_spool_double_clicked(self, sender, args):
        selected = self.line_numbers_lb.SelectedItem
        if not selected:
            self.set_status("Select a spool first.")
            return
        self.request_show(str(selected))

    def on_list_mouse_wheel(self, sender, args):
        parent = System.Windows.Media.VisualTreeHelper.GetParent(sender)
        while parent and not isinstance(parent, System.Windows.Controls.ScrollViewer):
            parent = System.Windows.Media.VisualTreeHelper.GetParent(parent)
        if parent:
            if args.Delta > 0:
                parent.LineUp()
            else:
                parent.LineDown()
            args.Handled = True

    def on_apply_clicked(self, sender, args):
        self.request_apply(self.spool_input_tb.Text, self.map_input_tb.Text)

    def on_rename_clicked(self, sender, args):
        selected = [str(item) for item in self.line_numbers_lb.SelectedItems]
        self.request_rename(self.find_tb.Text, self.replace_tb.Text, selected)

    def on_select_all_clicked(self, sender, args):
        self.line_numbers_lb.SelectAll()

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
        self._handler.map_name = None
        try:
            self._ext_event.Raise()
        except Exception:
            pass

    def request_show(self, spool_name):
        self._handler.request_name = "show"
        self._handler.spool_name = spool_name
        self._handler.map_name = None
        self._ext_event.Raise()

    def request_apply(self, spool_name, map_name):
        self._handler.request_name = "apply"
        self._handler.spool_name = spool_name
        self._handler.map_name = map_name
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
        self._handler.map_name = None
        self._ext_event.Raise()

    def request_tag_spools(self):
        self._handler.request_name = "tag_spools"
        self._handler.spool_name = None
        self._handler.map_name = None
        self._ext_event.Raise()

    def request_pin_all(self):
        self._handler.request_name = "pin_all"
        self._handler.spool_name = None
        self._handler.map_name = None
        self._ext_event.Raise()

    def set_all_spools(self, values):
        self._all_spools = list(values)
        self.apply_search_filter()

    def apply_search_filter(self):
        current_text = self.spool_input_tb.Text
        search_text = (self.search_tb.Text or "").strip().lower()

        if search_text:
            filtered = [x for x in self._all_spools if search_text in x.lower()]
        else:
            filtered = list(self._all_spools)

        self.line_numbers_lb.Items.Clear()
        for value in filtered:
            self.line_numbers_lb.Items.Add(value)

        self.spool_input_tb.Text = current_text

    def set_status(self, message):
        self.status_tb.Text = message or ""